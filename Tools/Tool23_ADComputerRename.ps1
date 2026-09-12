<#
.SYNOPSIS
    Tool 23: AD Remote Computer Rename & Attribute Cleanup
.DESCRIPTION
    - GUI-gestützte Remote-Umbenennung von Windows-Clients ohne WinRM.
    - Pfad-Auswahl, automatisches Erstellen und Öffnen der Quell-CSV (AlterName, NeuerName).
    - Prüft ICMP-Ping sowie SMB/RPC Port 445 vor dem Umbenennen.
    - Führt Rename-Computer mit -Force -Restart aus.
    - Verifiziert die Umbenennung im Active Directory (Polling bis 25 Sek.).
    - Passt AD-Attribute an: DisplayName, dNSHostName und Description-Auditlog.
    - Exportiert Fehlgeschlagene Computer automatisch als CSV.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

if (-not [System.Windows.Forms.Application]::RenderWithVisualStyles) {
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}
}

function Show-ADRenameComputersTool {
    $defaultDir = "C:\_Admin\Scripts\JHA"
    $defaultCsv = Join-Path $defaultDir "RenameClientsList.csv"
    $defaultLogDir = $defaultDir

    # Hauptformular
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Tool 23: AD Remote Computer Rename & Attribute Cleanup"
    $form.Size = New-Object System.Drawing.Size(1100, 750)
    $form.MinimumSize = New-Object System.Drawing.Size(950, 600)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # -------------------------------------------------------------------------
    # TOP PANEL: Pfade, CSV-Verwaltung & Start
    # -------------------------------------------------------------------------
    $pnlTop = New-Object System.Windows.Forms.GroupBox
    $pnlTop.Text = "Konfiguration & CSV-Steuerung"
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 115
    $pnlTop.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($pnlTop)

    # Zeile 1: Pfad zur CSV
    $lblCsv = New-Object System.Windows.Forms.Label
    $lblCsv.Text = "Rename-CSV Pfad:"
    $lblCsv.Location = New-Object System.Drawing.Point(15, 25)
    $lblCsv.Size = New-Object System.Drawing.Size(120, 22)
    $lblCsv.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Regular)
    $pnlTop.Controls.Add($lblCsv)

    $txtCsvPath = New-Object System.Windows.Forms.TextBox
    $txtCsvPath.Text = $defaultCsv
    $txtCsvPath.Location = New-Object System.Drawing.Point(140, 22)
    $txtCsvPath.Size = New-Object System.Drawing.Size(520, 23)
    $txtCsvPath.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $pnlTop.Controls.Add($txtCsvPath)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Durchsuchen..."
    $btnBrowse.Location = New-Object System.Drawing.Point(670, 20)
    $btnBrowse.Size = New-Object System.Drawing.Size(110, 27)
    $btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $pnlTop.Controls.Add($btnBrowse)

    $btnCreateTpl = New-Object System.Windows.Forms.Button
    $btnCreateTpl.Text = "Vorlage erstellen"
    $btnCreateTpl.Location = New-Object System.Drawing.Point(790, 20)
    $btnCreateTpl.Size = New-Object System.Drawing.Size(130, 27)
    $btnCreateTpl.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCreateTpl.BackColor = [System.Drawing.Color]::FromArgb(235, 240, 248)
    $pnlTop.Controls.Add($btnCreateTpl)

    $btnOpenCsv = New-Object System.Windows.Forms.Button
    $btnOpenCsv.Text = "CSV oeffnen / bearbeiten"
    $btnOpenCsv.Location = New-Object System.Drawing.Point(928, 20)
    $btnOpenCsv.Size = New-Object System.Drawing.Size(150, 27)
    $btnOpenCsv.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOpenCsv.BackColor = [System.Drawing.Color]::FromArgb(235, 240, 248)
    $pnlTop.Controls.Add($btnOpenCsv)

    # Zeile 2: Start-Button & Warnhinweis
    $btnStart = New-Object System.Windows.Forms.Button
    $btnStart.Text = "Umbenennung starten"
    $btnStart.Location = New-Object System.Drawing.Point(140, 62)
    $btnStart.Size = New-Object System.Drawing.Size(200, 36)
    $btnStart.BackColor = [System.Drawing.Color]::FromArgb(190, 40, 40)
    $btnStart.ForeColor = [System.Drawing.Color]::White
    $btnStart.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnStart.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($btnStart)

    $lblWarn = New-Object System.Windows.Forms.Label
    $lblWarn.Text = "Achtung: Fuehrt bei Erreichbarkeit sofortigen Neustart (-Force -Restart) des Zielclients aus!"
    $lblWarn.Location = New-Object System.Drawing.Point(355, 71)
    $lblWarn.AutoSize = $true
    $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(160, 40, 40)
    $lblWarn.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)
    $pnlTop.Controls.Add($lblWarn)

    # -------------------------------------------------------------------------
    # MITTE: SplitContainer mit Status-Grid (oben) und Log-Box (unten)
    # -------------------------------------------------------------------------
    $splitContainer = New-Object System.Windows.Forms.SplitContainer
    $splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitContainer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $splitContainer.SplitterDistance = 270
    $form.Controls.Add($splitContainer)
    $splitContainer.BringToFront()

    # Oben: Ergebnis-Tabelle
    $grpGrid = New-Object System.Windows.Forms.GroupBox
    $grpGrid.Text = "Verarbeitungsliste (Ergebnisse)"
    $grpGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpGrid.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $splitContainer.Panel1.Controls.Add($grpGrid)

    $gridResults = New-Object System.Windows.Forms.DataGridView
    $gridResults.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridResults.BackgroundColor = [System.Drawing.Color]::White
    $gridResults.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $gridResults.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $gridResults.ReadOnly = $true
    $gridResults.AllowUserToAddRows = $false
    $gridResults.AllowUserToDeleteRows = $false
    $gridResults.RowHeadersVisible = $false
    $gridResults.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $gridResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridResults.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Regular)
    $grpGrid.Controls.Add($gridResults)

    # Spalten definieren
    [void]$gridResults.Columns.Add("OldName", "Alter Name")
    [void]$gridResults.Columns.Add("NewName", "Neuer Name")
    [void]$gridResults.Columns.Add("Status", "Status")
    [void]$gridResults.Columns.Add("Details", "Details / Fehler")
    $gridResults.Columns["OldName"].FillWeight = 20
    $gridResults.Columns["NewName"].FillWeight = 20
    $gridResults.Columns["Status"].FillWeight = 25
    $gridResults.Columns["Details"].FillWeight = 35

    # Unten: Live-Logausgabe
    $grpLog = New-Object System.Windows.Forms.GroupBox
    $grpLog.Text = "Ausfuehrungs-Log"
    $grpLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $grpLog.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $splitContainer.Panel2.Controls.Add($grpLog)

    $txtLog = New-Object System.Windows.Forms.RichTextBox
    $txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtLog.BackColor = [System.Drawing.Color]::FromArgb(20, 24, 30)
    $txtLog.ForeColor = [System.Drawing.Color]::White
    $txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtLog.ReadOnly = $true
    $grpLog.Controls.Add($txtLog)

    # -------------------------------------------------------------------------
    # HILFSFUNKTIONEN
    # -------------------------------------------------------------------------
    $logAction = {
        param([string]$Msg, [System.Drawing.Color]$Color = [System.Drawing.Color]::White, [string]$FileLogPath = "")
        $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $entry = "[$timestamp] $Msg`r`n"

        $txtLog.SelectionStart = $txtLog.TextLength
        $txtLog.SelectionLength = 0
        $txtLog.SelectionColor = $Color
        $txtLog.AppendText($entry)
        $txtLog.ScrollToCaret()

        if ($FileLogPath) {
            try { $entry.TrimEnd("`r`n") | Out-File -FilePath $FileLogPath -Append -Encoding utf8 } catch {}
        }
    }

    # 1. Datei durchsuchen
    $btnBrowse.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "CSV-Dateien (*.csv)|*.csv|Alle Dateien (*.*)|*.*"
        $ofd.InitialDirectory = if (Test-Path $txtCsvPath.Text) { Split-Path $txtCsvPath.Text } else { "C:\" }
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtCsvPath.Text = $ofd.FileName
        }
    })

    # 2. Vorlage erzeugen
    $btnCreateTpl.Add_Click({
        $csvFile = $txtCsvPath.Text.Trim()
        $dir = Split-Path -Path $csvFile -Parent

        try {
            if (-not [string]::IsNullOrWhiteSpace($dir) -and (-not (Test-Path $dir))) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }

            if (-not (Test-Path $csvFile)) {
                $templateContent = @"
AlterName,NeuerName
CLIENT001,NB-GF-01
CLIENT002,WS-BU-05
"@
                $templateContent.Trim() | Out-File -FilePath $csvFile -Encoding utf8
                [System.Windows.Forms.MessageBox]::Show("Vorlage erfolgreich erstellt unter:`r`n$csvFile", "Erstellt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                [System.Windows.Forms.MessageBox]::Show("Die Datei existiert bereits unter:`r`n$csvFile", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Erstellen der Vorlage:`r`n$($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # 3. CSV öffnen
    $btnOpenCsv.Add_Click({
        $csvFile = $txtCsvPath.Text.Trim()
        if (-not (Test-Path $csvFile)) {
            $diag = [System.Windows.Forms.MessageBox]::Show("Die Datei existiert noch nicht. Soll sie jetzt neu erstellt werden?", "Datei nicht gefunden", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($diag -eq [System.Windows.Forms.DialogResult]::Yes) {
                $btnCreateTpl.PerformClick()
            } else {
                return
            }
        }
        if (Test-Path $csvFile) {
            Start-Process -FilePath $csvFile
        }
    })

    # 4. Umbenennungs-Vorgang starten
    $btnStart.Add_Click({
        $csvFile = $txtCsvPath.Text.Trim()
        if (-not (Test-Path $csvFile)) {
            [System.Windows.Forms.MessageBox]::Show("Die angegebene CSV-Datei existiert nicht!`r`n$csvFile", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # CSV prüfen
        try {
            $computers = @(Import-Csv -Path $csvFile -Delimiter ",")
            if ($computers.Count -eq 0 -or -not $computers[0].PSObject.Properties['AlterName'] -or -not $computers[0].PSObject.Properties['NeuerName']) {
                [System.Windows.Forms.MessageBox]::Show("Die CSV-Datei muss die Spalten 'AlterName' und 'NeuerName' (kommagetrennt) enthalten.", "Ungueltiges CSV-Format", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Einlesen der CSV:`r`n$($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Sicherheitsbestätigung
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Es wurden $($computers.Count) Computer in der CSV gefunden.`r`n`r`nSoll der Vorgang gestartet werden? Die Zielrechner werden nach erfolgreichem Befehl sofort neugestartet!",
            "Umbenennung bestaetigen",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        # Administrator-Credentials einholen
        $domainCred = $null
        try {
            $domainCred = Get-Credential -UserName "$env:USERDOMAIN\" -Message "Bitte Administrator-Zugangsdaten (mit Domain- und lokalen Admin-Rechten) eingeben"
        } catch { return }

        if ($null -eq $domainCred) { return }

        $btnStart.Enabled = $false
        $btnBrowse.Enabled = $false
        $btnCreateTpl.Enabled = $false
        $btnOpenCsv.Enabled = $false
        $gridResults.Rows.Clear()
        $txtLog.Clear()

        $adminUser = $domainCred.UserName
        $targetDir = Split-Path -Path $csvFile -Parent
        $logPath = Join-Path $targetDir "RenameLog_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').txt"
        $failCsvPath = Join-Path $targetDir "Fehlgeschlagene_Computer_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        $failedList = [System.Collections.Generic.List[PSCustomObject]]::new()

        & $logAction "=== START: Remote-Rename & AD-Attributbereinigung gestartet ===" ([System.Drawing.Color]::Cyan) $logPath
        & $logAction "Eintraege in Quell-CSV: $($computers.Count)" ([System.Drawing.Color]::White) $logPath

        # DNS-Root ermitteln
        try {
            $domainDns = (Get-ADDomain -Credential $domainCred).DNSRoot
        } catch {
            & $logAction "[-] Konnte Domaenen-DNS nicht per Get-ADDomain ermitteln. Fallback auf Host." ([System.Drawing.Color]::Orange) $logPath
            $domainDns = [System.Net.Dns]::GetHostByName("").HostName -replace '^[^.]*\.', ''
        }

        foreach ($entry in $computers) {
            $oldName = $entry.AlterName.Trim()
            $newName = $entry.NeuerName.Trim()
            $status = "Unbekannt"
            $fehlerDetails = ""

            & $logAction "--------------------------------------------------" ([System.Drawing.Color]::Gray) $logPath
            & $logAction "Verarbeite Client: $oldName -> $newName" ([System.Drawing.Color]::FromArgb(100, 180, 255)) $logPath
            [System.Windows.Forms.Application]::DoEvents()

            # Schritt 1: Ping-Prüfung
            if (-not (Test-Connection -ComputerName $oldName -Count 1 -Quiet)) {
                $status = "Offline"
                $fehlerDetails = "Client antwortet nicht auf ICMP (Ping)."
                & $logAction "[-] $fehlerDetails" ([System.Drawing.Color]::Orange) $logPath
            } else {
                & $logAction "[+] Ping erfolgreich." ([System.Drawing.Color]::LightGray) $logPath

                # Schritt 2: Port 445 Prüfung
                $tcpSocket = New-Object System.Net.Sockets.TcpClient
                $portOk = $false
                try {
                    $asyncResult = $tcpSocket.BeginConnect($oldName, 445, $null, $null)
                    $waitHandle = $asyncResult.AsyncWaitHandle.WaitOne(1500, $false)
                    if ($waitHandle -and $tcpSocket.Connected) {
                        $tcpSocket.EndConnect($asyncResult)
                        $portOk = $true
                    }
                } catch {
                    $portOk = $false
                } finally {
                    $tcpSocket.Close()
                    $tcpSocket.Dispose()
                }

                if (-not $portOk) {
                    $status = "RPC/SMB blockiert"
                    $fehlerDetails = "Port 445 nicht erreichbar (Datei- und Druckerfreigabe / Client-Firewall)."
                    & $logAction "[-] $fehlerDetails" ([System.Drawing.Color]::Orange) $logPath
                } else {
                    & $logAction "[+] Port 445 (SMB/RPC) erreichbar. Sende Rename-Befehl..." ([System.Drawing.Color]::LightGray) $logPath

                    # Schritt 3: Rename-Computer
                    try {
                        Rename-Computer -ComputerName $oldName -NewName $newName -DomainCredential $domainCred -Force -Restart -ErrorAction Stop
                        & $logAction "[*] Befehl gesendet. Warte auf AD-Synchronisation..." ([System.Drawing.Color]::Yellow) $logPath
                        [System.Windows.Forms.Application]::DoEvents()

                        # Schritt 4: AD-Verifikation
                        $adCheck = $null
                        $timeoutSeconds = 25
                        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

                        while ($stopwatch.Elapsed.TotalSeconds -lt $timeoutSeconds) {
                            Start-Sleep -Seconds 3
                            $adCheck = Get-ADComputer -Filter "Name -eq '$newName' -or sAMAccountName -eq '$newName$'" -Credential $domainCred -Properties Description, DisplayName, dNSHostName -ErrorAction SilentlyContinue
                            if ($adCheck) { break }
                            [System.Windows.Forms.Application]::DoEvents()
                        }
                        $stopwatch.Stop()

                        if ($adCheck) {
                            & $logAction "[SUCCESS] $oldName erfolgreich in $newName umbenannt (AD bestaetigt)!" ([System.Drawing.Color]::LightGreen) $logPath

                            # Schritt 5: AD-Attribute anpassen
                            try {
                                $expectedDns = "$newName.$domainDns".ToLower()
                                $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                $descEntry = "Umbenannt von $oldName auf $newName am $ts durch $adminUser"

                                $newDesc = if ([string]::IsNullOrWhiteSpace($adCheck.Description)) {
                                    $descEntry
                                } else {
                                    "$($adCheck.Description) | $descEntry"
                                }

                                $setAttr = @{
                                    DisplayName = $newName
                                    Description = $newDesc
                                }

                                if ($adCheck.dNSHostName -ne $expectedDns) {
                                    $setAttr["dNSHostName"] = $expectedDns
                                    & $logAction "[*] dNSHostName korrigiert: $expectedDns" ([System.Drawing.Color]::Yellow) $logPath
                                }

                                Set-ADComputer -Identity $adCheck.DistinguishedName -Credential $domainCred @setAttr -ErrorAction Stop
                                & $logAction "[+] AD-Attribute angepasst (DisplayName, Description)." ([System.Drawing.Color]::LightGreen) $logPath
                                $status = "Erfolgreich"
                            } catch {
                                $status = "Erfolgreich mit Attributwarnung"
                                $fehlerDetails = "PC umbenannt, Attribute unvollstaendig: $($_.Exception.Message)"
                                & $logAction "[!] $fehlerDetails" ([System.Drawing.Color]::Yellow) $logPath
                            }
                        } else {
                            $status = "AD-Verifikation Timeout"
                            $fehlerDetails = "Rename abgesetzt, Objekt nach 25s im AD noch nicht als '$newName' auffindbar."
                            & $logAction "[-] $fehlerDetails" ([System.Drawing.Color]::Salmon) $logPath
                        }
                    } catch {
                        $status = "Rename fehlgeschlagen"
                        $fehlerDetails = $_.Exception.Message
                        & $logAction "[-] Rename fehlgeschlagen: $fehlerDetails" ([System.Drawing.Color]::Salmon) $logPath
                    }
                }
            }

            # In GUI-Tabelle eintragen
            $rowIndex = $gridResults.Rows.Add($oldName, $newName, $status, $fehlerDetails)
            $row = $gridResults.Rows[$rowIndex]
            if ($status -eq "Erfolgreich") {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 250, 235)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
            } else {
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 238, 238)
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
                $failedList.Add([PSCustomObject]@{
                    AlterName        = $oldName
                    GewuenschterName = $newName
                    Status           = $status
                    FehlerDetails    = $fehlerDetails
                })
            }
            [System.Windows.Forms.Application]::DoEvents()
        }

        # Abschluss & Fehlerdatei
        & $logAction "--------------------------------------------------" ([System.Drawing.Color]::Gray) $logPath
        if ($failedList.Count -gt 0) {
            try {
                $failedList | Export-Csv -Path $failCsvPath -NoTypeInformation -Delimiter ";" -Encoding utf8
                & $logAction "=== FINISH: Beendet mit $($failedList.Count) Warnung(en)/Fehlern. ===" ([System.Drawing.Color]::Orange) $logPath
                & $logAction "Fehlerliste exportiert: $failCsvPath" ([System.Drawing.Color]::Orange) $logPath
            } catch {}
            [System.Windows.Forms.MessageBox]::Show("Der Vorgang wurde mit $($failedList.Count) Warnung(en)/Fehler(n) beendet.`r`nFehlerbericht:`r`n$failCsvPath", "Abgeschlossen mit Fehlern", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        } else {
            & $logAction "=== FINISH: Alle Clients erfolgreich verarbeitet! ===" ([System.Drawing.Color]::LightGreen) $logPath
            [System.Windows.Forms.MessageBox]::Show("Alle Computer wurden erfolgreich umbenannt und im AD aktualisiert.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }

        $btnStart.Enabled = $true
        $btnBrowse.Enabled = $true
        $btnCreateTpl.Enabled = $true
        $btnOpenCsv.Enabled = $true
    })

    [void]$form.ShowDialog()
}

# Standalone Aufruf
Show-ADRenameComputersTool
