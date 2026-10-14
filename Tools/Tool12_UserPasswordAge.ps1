<#
================================================================================
 TOOL 12: AD BENUTZER KENNWORTALTER & MULTI-DC LASTLOGON AUDIT
 Inklusive CSV-Export (mit Pfadauswahl), Snapshot-Export und Vergleich zweier Stände.
 100% abgesichert mit Semikolons und methodenbasierter Filterung (keine Pipes).
================================================================================
#>

Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue;
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue;

function Open-ToolUserPasswordAge {
    if (Get-Command Assert-DomainJoined -ErrorAction SilentlyContinue) {
        if (-not (Assert-DomainJoined)) { return; }
    }

    $isDE = ($script:CurrentLang -eq "DE");

    $subForm = New-Object System.Windows.Forms.Form;
    $subForm.Text = if ($isDE) { "Tool 12: AD Benutzer Kennwortalter & Multi-DC LastLogon" } else { "Tool 12: AD User Password Age & Multi-DC LastLogon" };
    $subForm.Size = New-Object System.Drawing.Size(1250, 800);
    $subForm.StartPosition = "CenterParent";
    $subForm.Font = New-Object System.Drawing.Font("Segoe UI", 9);
    $subForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250);

    # 1. TOP PANEL
    $pnlTop = New-Object System.Windows.Forms.Panel;
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top;
    $pnlTop.Height = 85;
    $pnlTop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250);
    $subForm.Controls.Add($pnlTop);

    $lblSearch = New-Object System.Windows.Forms.Label;
    $lblSearch.Location = New-Object System.Drawing.Point(15, 14);
    $lblSearch.Text = if ($isDE) { "Benutzer (Wildcard):" } else { "User (Wildcard):" };
    $lblSearch.Size = New-Object System.Drawing.Size(130, 20);
    $pnlTop.Controls.Add($lblSearch);

    $txtSearch = New-Object System.Windows.Forms.TextBox;
    $txtSearch.Location = New-Object System.Drawing.Point(150, 11);
    $txtSearch.Size = New-Object System.Drawing.Size(170, 23);
    $txtSearch.Text = "*";
    $pnlTop.Controls.Add($txtSearch);

    $lblSort = New-Object System.Windows.Forms.Label;
    $lblSort.Location = New-Object System.Drawing.Point(335, 14);
    $lblSort.Text = if ($isDE) { "Sortieren nach:" } else { "Sort by:" };
    $lblSort.Size = New-Object System.Drawing.Size(95, 20);
    $pnlTop.Controls.Add($lblSort);

    $cmbSort = New-Object System.Windows.Forms.ComboBox;
    $cmbSort.Location = New-Object System.Drawing.Point(435, 11);
    $cmbSort.Size = New-Object System.Drawing.Size(145, 23);
    $cmbSort.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $cmbSort.Items.AddRange(@("Benutzername", "Anzeigename", "Status", "PW-Alter (Tage)", "Neuester Login"));
    $cmbSort.SelectedIndex = 0;
    $pnlTop.Controls.Add($cmbSort);

    $lblStatusFilter = New-Object System.Windows.Forms.Label;
    $lblStatusFilter.Location = New-Object System.Drawing.Point(15, 48);
    $lblStatusFilter.Text = if ($isDE) { "Konto-Status:" } else { "Account Status:" };
    $lblStatusFilter.Size = New-Object System.Drawing.Size(130, 20);
    $pnlTop.Controls.Add($lblStatusFilter);

    $cmbStatusFilter = New-Object System.Windows.Forms.ComboBox;
    $cmbStatusFilter.Location = New-Object System.Drawing.Point(150, 45);
    $cmbStatusFilter.Size = New-Object System.Drawing.Size(170, 23);
    $cmbStatusFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $cmbStatusFilter.Items.AddRange(@("Alle Status", "Nur Aktiv", "Nur Deaktiviert"));
    $cmbStatusFilter.SelectedIndex = 0;
    $pnlTop.Controls.Add($cmbStatusFilter);

    # Buttons
    $btnRun = New-Object System.Windows.Forms.Button;
    $btnRun.Location = New-Object System.Drawing.Point(595, 11);
    $btnRun.Size = New-Object System.Drawing.Size(100, 58);
    $btnRun.Text = if ($isDE) { "Abfragen" } else { "Query" };
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215);
    $btnRun.ForeColor = [System.Drawing.Color]::White;
    $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat;
    $btnRun.FlatAppearance.BorderSize = 0;
    $btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold);
    $pnlTop.Controls.Add($btnRun);

    $btnExportCsv = New-Object System.Windows.Forms.Button;
    $btnExportCsv.Location = New-Object System.Drawing.Point(705, 11);
    $btnExportCsv.Size = New-Object System.Drawing.Size(100, 27);
    $btnExportCsv.Text = if ($isDE) { "CSV Export..." } else { "Export CSV..." };
    $btnExportCsv.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230);
    $btnExportCsv.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat;
    $pnlTop.Controls.Add($btnExportCsv);

    $btnSaveSnap = New-Object System.Windows.Forms.Button;
    $btnSaveSnap.Location = New-Object System.Drawing.Point(705, 42);
    $btnSaveSnap.Size = New-Object System.Drawing.Size(100, 27);
    $btnSaveSnap.Text = if ($isDE) { "Snapshot..." } else { "Save Snapshot" };
    $btnSaveSnap.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 245);
    $btnSaveSnap.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat;
    $pnlTop.Controls.Add($btnSaveSnap);

    $btnCompare = New-Object System.Windows.Forms.Button;
    $btnCompare.Location = New-Object System.Drawing.Point(815, 11);
    $btnCompare.Size = New-Object System.Drawing.Size(145, 58);
    $btnCompare.Text = if ($isDE) { "Snapshots`nvergleichen" } else { "Compare`nSnapshots" };
    $btnCompare.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 230);
    $btnCompare.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat;
    $btnCompare.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold);
    $pnlTop.Controls.Add($btnCompare);

    $lblStatus = New-Object System.Windows.Forms.Label;
    $lblStatus.Location = New-Object System.Drawing.Point(970, 14);
    $lblStatus.Size = New-Object System.Drawing.Size(265, 58);
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(30, 58, 138);
    $lblStatus.Text = if ($isDE) { "Lade Domain Controller Sites..." } else { "Loading Domain Controller Sites..." };
    $pnlTop.Controls.Add($lblStatus);

    # 2. DC PANEL
    $grpDCs = New-Object System.Windows.Forms.GroupBox;
    $grpDCs.Text = if ($isDE) { "Verfuegbare Domain Controller (Waehlen Sie die abzufragenden DCs aus)" } else { "Available Domain Controllers (Select DCs to query)" };
    $grpDCs.Dock = [System.Windows.Forms.DockStyle]::Top;
    $grpDCs.Height = 140;
    $grpDCs.Font = New-Object System.Drawing.Font("Segoe UI", 8.5);
    $subForm.Controls.Add($grpDCs);

    $gridDCs = New-Object System.Windows.Forms.DataGridView;
    $gridDCs.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $gridDCs.AllowUserToAddRows = $false;
    $gridDCs.AllowUserToDeleteRows = $false;
    $gridDCs.RowHeadersVisible = $false;
    $gridDCs.BackgroundColor = [System.Drawing.Color]::White;
    $gridDCs.BorderStyle = [System.Windows.Forms.BorderStyle]::None;
    $gridDCs.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill;
    $grpDCs.Controls.Add($gridDCs);

    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn;
    $colCheck.Name = "Check";
    $colCheck.HeaderText = if ($isDE) { "Abfragen" } else { "Query" };
    $colCheck.Width = 70;
    [void]$gridDCs.Columns.Add($colCheck);
    [void]$gridDCs.Columns.Add("HostName", "Domain Controller");
    [void]$gridDCs.Columns.Add("Site", "AD Site");
    $gridDCs.Columns[0].FillWeight = 15;
    $gridDCs.Columns[1].FillWeight = 55;
    $gridDCs.Columns[2].FillWeight = 30;

    # 3. STATUS-LEGENDE
    $pnlLegend = New-Object System.Windows.Forms.Panel;
    $pnlLegend.Dock = [System.Windows.Forms.DockStyle]::Bottom;
    $pnlLegend.Height = 38;
    $pnlLegend.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250);
    $subForm.Controls.Add($pnlLegend);

    $lblLegTitle = New-Object System.Windows.Forms.Label;
    $lblLegTitle.Location = New-Object System.Drawing.Point(15, 11);
    $lblLegTitle.Text = "Status-Legende:";
    $lblLegTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold);
    $lblLegTitle.AutoSize = $true;
    $pnlLegend.Controls.Add($lblLegTitle);

    $pnlBoxGreen = New-Object System.Windows.Forms.Panel;
    $pnlBoxGreen.Location = New-Object System.Drawing.Point(115, 12);
    $pnlBoxGreen.Size = New-Object System.Drawing.Size(12, 12);
    $pnlBoxGreen.BackColor = [System.Drawing.Color]::FromArgb(220, 252, 231);
    $pnlBoxGreen.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle;
    $pnlLegend.Controls.Add($pnlBoxGreen);

    $lblLegGreen = New-Object System.Windows.Forms.Label;
    $lblLegGreen.Location = New-Object System.Drawing.Point(132, 11);
    $lblLegGreen.Text = "Login <= 90 Tage";
    $lblLegGreen.AutoSize = $true;
    $pnlLegend.Controls.Add($lblLegGreen);

    $pnlBoxRed = New-Object System.Windows.Forms.Panel;
    $pnlBoxRed.Location = New-Object System.Drawing.Point(245, 12);
    $pnlBoxRed.Size = New-Object System.Drawing.Size(12, 12);
    $pnlBoxRed.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226);
    $pnlBoxRed.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle;
    $pnlLegend.Controls.Add($pnlBoxRed);

    $lblLegRed = New-Object System.Windows.Forms.Label;
    $lblLegRed.Location = New-Object System.Drawing.Point(262, 11);
    $lblLegRed.Text = "Login > 90 Tage / Nie";
    $lblLegRed.AutoSize = $true;
    $pnlLegend.Controls.Add($lblLegRed);

    $lblLegInfo = New-Object System.Windows.Forms.Label;
    $lblLegInfo.Location = New-Object System.Drawing.Point(415, 11);
    $lblLegInfo.Text = "| Grau = Konto Deaktiviert | Klick auf Spaltenkopf sortiert die Tabelle";
    $lblLegInfo.ForeColor = [System.Drawing.Color]::FromArgb(100, 116, 139);
    $lblLegInfo.AutoSize = $true;
    $pnlLegend.Controls.Add($lblLegInfo);

    # 4. RESULT GRID
    $gridResults = New-Object System.Windows.Forms.DataGridView;
    $gridResults.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $gridResults.ReadOnly = $true;
    $gridResults.AllowUserToAddRows = $false;
    $gridResults.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect;
    $gridResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells;
    $gridResults.BackgroundColor = [System.Drawing.Color]::White;
    $gridResults.BorderStyle = [System.Windows.Forms.BorderStyle]::None;
    $gridResults.RowHeadersVisible = $false;
    $gridResults.EnableHeadersVisualStyles = $false;
    $gridResults.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249);
    $gridResults.ColumnHeadersHeight = 32;
    $subForm.Controls.Add($gridResults);

    $gridResults.BringToFront();
    $pnlLegend.SendToBack();
    $grpDCs.SendToBack();
    $pnlTop.SendToBack();

    # HILFSFUNKTIONEN
    $testDcReachability = {
        param([string]$dcHost, [int]$timeoutMs = 3000);
        try {
            $tcpClient = New-Object System.Net.Sockets.TcpClient;
            $asyncResult = $tcpClient.BeginConnect($dcHost, 389, $null, $null);
            $success = $asyncResult.AsyncWaitHandle.WaitOne($timeoutMs, $false);
            if ($success -and $tcpClient.Connected) {
                $tcpClient.EndConnect($asyncResult);
                $tcpClient.Close();
                return $true;
            } else {
                $tcpClient.Close();
                return $false;
            }
        } catch { return $false; }
    };

    $convertLargeInt = {
        param($val);
        if ($null -eq $val) { return [int64]0; }
        if ($val -is [int64]) { return $val; }
        if ($val -is [int32] -or $val -is [int] -or $val -is [double]) { return [int64]$val; }
        try {
            $type = $val.GetType();
            $high = [int64]$type.InvokeMember("HighPart", [System.Reflection.BindingFlags]::GetProperty, $null, $val, $null);
            $low  = [int64]$type.InvokeMember("LowPart", [System.Reflection.BindingFlags]::GetProperty, $null, $val, $null);
            if ($low -lt 0) { $low = [int64]([uint32]$low); }
            return [int64](($high -shl 32) + $low);
        } catch {}
        try {
            $high = [int64]$val.HighPart;
            $low  = [int64]$val.LowPart;
            if ($low -lt 0) { $low = [int64]([uint32]$low); }
            return [int64](($high -shl 32) + $low);
        } catch {}
        try {
            [int64]$parsed = 0;
            if ([int64]::TryParse($val.ToString(), [ref]$parsed)) { return $parsed; }
        } catch {}
        return [int64]0;
    };

    $formatFileTime = {
        param([int64]$ticks);
        if ($ticks -le 0 -or $ticks -ge [DateTime]::MaxValue.ToFileTime()) { return $null; }
        try { return [DateTime]::FromFileTime($ticks); } catch { return $null; }
    };

    $script:rawUserResults = @();

    $applyRowColors = {
        $ninetyDaysAgo = (Get-Date).AddDays(-90);
        for ($rIdx = 0; $rIdx -lt $gridResults.Rows.Count; $rIdx++) {
            $row = $gridResults.Rows[$rIdx];
            $statusVal = $row.Cells["Status"].Value;
            $lastLoginVal = $row.Cells["Neuester Login"].Value;

            if ($statusVal -eq "Deaktiviert") {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray;
                $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White;
            } else {
                $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black;
                if ($lastLoginVal -and $lastLoginVal -ne "Nie" -and $lastLoginVal -ne "Offline / Nicht erreichbar") {
                    try {
                        $parsed = [datetime]::ParseExact($lastLoginVal, "dd.MM.yyyy HH:mm", $null);
                        if ($parsed -ge $ninetyDaysAgo) {
                            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(220, 252, 231);
                        } else {
                            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226);
                        }
                    } catch {}
                } else {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226);
                }
            }
        }
    };

    $applyFilterAndSort = {
        if (-not $script:rawUserResults -or $script:rawUserResults.Count -eq 0) { return; }

        $statusSel = $cmbStatusFilter.SelectedIndex;
        $sortProp  = $cmbSort.SelectedItem.ToString();

        $filtered = @($script:rawUserResults).Where({
            switch ($statusSel) {
                1 { $_.Status -eq "Aktiv"; }
                2 { $_.Status -eq "Deaktiviert"; }
                Default { $true; }
            }
        });

        $sorted = [System.Collections.ArrayList]::new();
        $sortObj = $filtered | Sort-Object {
            if ($sortProp -eq "PW-Alter (Tage)") {
                $v = $_.$sortProp;
                if ($v -is [int]) { $v; } else { 99999; }
            } else {
                $_.$sortProp;
            }
        };

        foreach ($item in $sortObj) {
            [void]$sorted.Add($item);
        }

        $gridResults.DataSource = $sorted;
        & $applyRowColors;
        $lblStatus.Text = "Fertig: $($sorted.Count) von $($script:rawUserResults.Count) Konten angezeigt.";
    };

    # INITIALES LADEN DER DCS
    $subForm.Add_Shown({
        $lblStatus.Text = "Lade Domain Controller Sites...";
        $subForm.Refresh();
        $gridDCs.Rows.Clear();

        try {
            $dom = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain();
            $dcList = @($dom.DomainControllers);
            for ($d = 0; $d -lt $dcList.Count; $d++) {
                $dc = $dcList[$d];
                $sName = if ($dc.SiteName) { $dc.SiteName; } else { "Unbekannt"; };
                [void]$gridDCs.Rows.Add($true, $dc.Name, $sName);
            }
            $lblStatus.Text = "Bereit. $($dom.DomainControllers.Count) DCs gefunden.";
        } catch {
            $fallback = if ($env:LOGONSERVER) { $env:LOGONSERVER.Replace("\", ""); } else { "localhost"; };
            [void]$gridDCs.Rows.Add($true, $fallback, "Default-Site");
            $lblStatus.Text = "Bereit (Fallback DC).";
        }
    });

    # ABFRAGE DURCHFÜHREN
    $btnRun.Add_Click({
        $subForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor;
        $lblStatus.Text = "Pruefe DCs...";
        [System.Windows.Forms.Application]::DoEvents();

        $activeDCs = @();
        for ($g = 0; $g -lt $gridDCs.Rows.Count; $g++) {
            $row = $gridDCs.Rows[$g];
            if ($row.Cells["Check"].Value -eq $true) {
                $dcHost = $row.Cells["HostName"].Value;
                $lblStatus.Text = "Ping $dcHost...";
                [System.Windows.Forms.Application]::DoEvents();

                $isOnline = & $testDcReachability $dcHost 3000;
                if ($isOnline) {
                    $activeDCs += $dcHost;
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::White;
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black;
                } else {
                    $row.Cells["Check"].Value = $false;
                    $row.Cells["Site"].Value = "Timeout (> 3s)";
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226);
                    $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed;
                }
            }
        }

        if ($activeDCs.Count -eq 0) {
            $subForm.Cursor = [System.Windows.Forms.Cursors]::Default;
            [System.Windows.Forms.MessageBox]::Show("Kein DC erreichbar.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
            return;
        }

        $lblStatus.Text = "AD-Suche laeuft...";
        [System.Windows.Forms.Application]::DoEvents();

        try {
            $rootEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://RootDSE");
            $defNC = $rootEntry.defaultNamingContext.ToString();
            $now = Get-Date;

            $searchPattern = $txtSearch.Text.Trim();
            if ([string]::IsNullOrEmpty($searchPattern)) { $searchPattern = "*"; }

            $ldapFilter = if ($searchPattern -eq "*") {
                "(&(objectCategory=person)(objectClass=user))";
            } else {
                "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=$searchPattern)(cn=$searchPattern)))";
            };

            $primaryDC = $activeDCs[0];
            $domEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$primaryDC/$defNC");
            $searcher = [System.DirectoryServices.DirectorySearcher]::new($domEntry);
            $searcher.Filter = $ldapFilter;
            $searcher.PageSize = 2500;
            $searcher.PropertiesToLoad.AddRange(@("samaccountname", "displayname", "pwdlastset", "useraccountcontrol", "distinguishedname"));
            $allUsers = $searcher.FindAll();

            if (-not $allUsers -or $allUsers.Count -eq 0) {
                $lblStatus.Text = "Keine Benutzer gefunden.";
                $gridResults.DataSource = $null;
                return;
            }

            $dcLoginData = @{};
            for ($dcIdx = 0; $dcIdx -lt $activeDCs.Count; $dcIdx++) {
                $dc = $activeDCs[$dcIdx];
                $dcLoginData[$dc] = @{};
                $lblStatus.Text = "Lese LastLogon von DC: $dc...";
                [System.Windows.Forms.Application]::DoEvents();

                try {
                    $dcEntry = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$dc/$defNC");
                    $dcSearcher = [System.DirectoryServices.DirectorySearcher]::new($dcEntry);
                    $dcSearcher.Filter = $ldapFilter;
                    $dcSearcher.PageSize = 2500;
                    $dcSearcher.PropertiesToLoad.AddRange(@("samaccountname", "lastlogon"));
                    $dcResults = $dcSearcher.FindAll();

                    for ($resIdx = 0; $resIdx -lt $dcResults.Count; $resIdx++) {
                        $res = $dcResults[$resIdx];
                        $sName = if ($res.Properties["samaccountname"].Count -gt 0) { $res.Properties["samaccountname"][0].ToString().ToLower(); } else { $null; };
                        if ($sName) {
                            $rawLL = if ($res.Properties["lastlogon"].Count -gt 0) { $res.Properties["lastlogon"][0]; } else { 0; };
                            $dcLoginData[$dc][$sName] = & $convertLargeInt $rawLL;
                        }
                    }
                } catch {}
            }

            $resultsList = [System.Collections.ArrayList]::new();
            for ($uIdx = 0; $uIdx -lt $allUsers.Count; $uIdx++) {
                $u = $allUsers[$uIdx];
                $p = $u.Properties;
                $sam = if ($p["samaccountname"].Count -gt 0) { $p["samaccountname"][0].ToString(); } else { ""; };
                $disp = if ($p["displayname"].Count -gt 0) { $p["displayname"][0].ToString(); } else { ""; };
                $dn = if ($p["distinguishedname"].Count -gt 0) { $p["distinguishedname"][0].ToString(); } else { ""; };

                $ou = "Users";
                if ($dn -match 'OU=([^,]+)') { $ou = $Matches[1]; }

                $uac = if ($p["useraccountcontrol"].Count -gt 0) { [int]$p["useraccountcontrol"][0]; } else { 0; };
                $isDisabled = ($uac -band 2) -eq 2;
                $dontExpire = ($uac -band 65536) -eq 65536;

                $rawPwd = if ($p["pwdlastset"].Count -gt 0) { $p["pwdlastset"][0]; } else { 0; };
                $pwdTicks = & $convertLargeInt $rawPwd;
                $pwdDateStr = "";
                $pwdAgeDays = 9999;

                if ($pwdTicks -eq 0) {
                    $pwdDateStr = "Muss geaendert werden (pwdLastSet=0)";
                } else {
                    $dtPwd = & $formatFileTime $pwdTicks;
                    if ($null -ne $dtPwd) {
                        $pwdDateStr = $dtPwd.ToString("dd.MM.yyyy HH:mm");
                        $pwdAgeDays = [int][math]::Round(($now - $dtPwd).TotalDays, 0);
                    } else {
                        $pwdDateStr = "Nie";
                    }
                }

                $samLower = $sam.ToLower();
                $maxLogonTicks = 0;
                $latestDC = "N/A";
                $dcValues = [ordered]@{};

                for ($r = 0; $r -lt $gridDCs.Rows.Count; $r++) {
                    $dc = $gridDCs.Rows[$r].Cells["HostName"].Value;
                    if ($activeDCs -contains $dc) {
                        $llTicks = 0;
                        if ($dcLoginData.ContainsKey($dc) -and $dcLoginData[$dc].ContainsKey($samLower)) {
                            $llTicks = $dcLoginData[$dc][$samLower];
                        }
                        $dtLogon = & $formatFileTime $llTicks;
                        if ($null -ne $dtLogon) {
                            $dcValues[$dc] = $dtLogon.ToString("dd.MM.yyyy HH:mm");
                            if ($llTicks -gt $maxLogonTicks) {
                                $maxLogonTicks = $llTicks;
                                $latestDC = $dc;
                            }
                        } else {
                            $dcValues[$dc] = "Nie";
                        }
                    } else {
                        $dcValues[$dc] = "Offline / Nicht erreichbar";
                    }
                }

                $latestLoginStr = "Nie";
                $inactDays = 9999;
                if ($maxLogonTicks -gt 0) {
                    $dtMax = & $formatFileTime $maxLogonTicks;
                    if ($null -ne $dtMax) {
                        $latestLoginStr = $dtMax.ToString("dd.MM.yyyy HH:mm");
                        $inactDays = [int][math]::Round(($now - $dtMax).TotalDays, 0);
                    }
                }

                $rowMap = [ordered]@{
                    "Benutzername"             = $sam;
                    "Anzeigename"              = $disp;
                    "Status"                   = if ($isDisabled) { "Deaktiviert"; } else { "Aktiv"; };
                    "Letzte Kennwortaenderung" = $pwdDateStr;
                    "PW-Alter (Tage)"          = if ($pwdDateStr -eq "Muss geaendert werden (pwdLastSet=0)") { 0; } else { $pwdAgeDays; };
                    "PW laeuft nie ab"         = if ($dontExpire) { "Ja"; } else { "Nein"; };
                    "Neuester Login"           = $latestLoginStr;
                    "Neuester DC"              = $latestDC;
                    "Inaktiv seit (Tage)"      = if ($latestLoginStr -eq "Nie") { "Nie"; } else { $inactDays; };
                    "OU Pfad"                  = $ou;
                };

                $dcKeyList = @($dcValues.Keys);
                for ($k = 0; $k -lt $dcKeyList.Count; $k++) {
                    $keyName = $dcKeyList[$k];
                    $rowMap["DC: $keyName"] = $dcValues[$keyName];
                }

                [void]$resultsList.Add([PSCustomObject]$rowMap);
            }

            $script:rawUserResults = $resultsList;
            & $applyFilterAndSort;

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Scan: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
            $lblStatus.Text = "Fehler: " + $_.Exception.Message;
        } finally {
            $subForm.Cursor = [System.Windows.Forms.Cursors]::Default;
        }
    });

    # EXPORT: CSV MIT PFADAUSWAHL
    $btnExportCsv.Add_Click({
        if (-not $gridResults.DataSource -or $gridResults.Rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information);
            return;
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog;
        $sfd.Title = "Benutzer-Kennwortalter & Logins exportieren";
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv|Alle Dateien (*.*)|*.*";
        $sfd.FileName = "AD_User_PasswordAge_Audit_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv";

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $exportData = @($gridResults.DataSource);
                $exportData | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";";
                [System.Windows.Forms.MessageBox]::Show("Export erfolgreich gespeichert:`n$($sfd.FileName)", "Export abgeschlossen", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information);
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Export: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
            }
        }
    });

    # EXPORT: SNAPSHOT SPEICHERN
    $btnSaveSnap.Add_Click({
        if (-not $script:rawUserResults -or $script:rawUserResults.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte fuehren Sie zuerst einen Scan aus.", "Snapshot", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning);
            return;
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog;
        $sfd.Title = "AD User Audit Snapshot speichern";
        $sfd.Filter = "Snapshot Datei (*.pwsnap)|*.pwsnap|XML Datei (*.xml)|*.xml";
        $sfd.FileName = "AD_User_Snapshot_$((Get-Date).ToString('yyyyMMdd_HHmm')).pwsnap";

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $snapData = @{
                    Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss");
                    Domain    = $env:USERDNSDOMAIN;
                    Data      = @($script:rawUserResults);
                };
                $snapData | Export-Clixml -Path $sfd.FileName -Depth 4;
                [System.Windows.Forms.MessageBox]::Show("Snapshot erfolgreich gesichert!`n$($sfd.FileName)", "Snapshot erstellt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information);
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
            }
        }
    });

    # FUNKTION: VERGLEICH ZWEIER SNAPSHOTS
    $btnCompare.Add_Click({$ofd1 = New-Object System.Windows.Forms.OpenFileDialog;
        $ofd1.Title = "Snapshot 1 (AELTERER Stand) auswaehlen";
        $ofd1.Filter = "Snapshot Datei (*.pwsnap;*.xml)|*.pwsnap;*.xml|Alle Dateien (*.*)|*.*";
        if ($ofd1.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return; }

        $ofd2 = New-Object System.Windows.Forms.OpenFileDialog;
        $ofd2.Title = "Snapshot 2 (NEUERER Stand) auswaehlen";
        $ofd2.Filter = "Snapshot Datei (*.pwsnap;*.xml)|*.pwsnap;*.xml|Alle Dateien (*.*)|*.*";
        if ($ofd2.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return; }

        try {
            $snap1 = Import-Clixml -Path$ofd1.FileName;
            $snap2 = Import-Clixml -Path$ofd2.FileName;

            $list1 = if ($snap1.Data) { $snap1.Data; } else {$snap1; };
            $list2 = if ($snap2.Data) { $snap2.Data; } else {$snap2; };

            $map1 = @{};
            for ($i = 0; $i -lt$list1.Count; $i++) {$u = $list1[$i];
                if ($u.Benutzername) {$map1[$u.Benutzername.ToLower()] =$u; }
            }

            $map2 = @{};
            for ($j = 0; $j -lt$list2.Count; $j++) {$u = $list2[$j];
                if ($u.Benutzername) {$map2[$u.Benutzername.ToLower()] =$u; }
            }

            $allUsers = [System.Collections.Generic.HashSet[string]]::new();
            $kList1 = @($map1.Keys);
            for ($k1 = 0; $k1 -lt$kList1.Count; $k1++) { [void]$allUsers.Add($kList1[$k1]); }
            $kList2 = @($map2.Keys);
            for ($k2 = 0; $k2 -lt$kList2.Count; $k2++) { [void]$allUsers.Add($kList2[$k2]); }

            $diffRows = [System.Collections.ArrayList]::new();$sortedKeys = [System.Collections.ArrayList]::new();
            $keysSort =$allUsers | Sort-Object;
            foreach ($sk in $keysSort) { [void]$sortedKeys.Add($sk); }

            for ($idx = 0; $idx -lt$sortedKeys.Count; $idx++) {$k = $sortedKeys[$idx];
                $u1 =$map1[$k];$u2 = $map2[$k];

                $sam   = if ($u2) {$u2.Benutzername; } else { $u1.Benutzername; };$disp  = if ($u2) {$u2.Anzeigename; } else { $u1.Anzeigename; };$ou    = if ($u2) {$u2."OU Pfad"; } else { $u1."OU Pfad"; };

                $status1 = if ($u1) {$u1.Status; } else { "-"; };
                $status2 = if ($u2) {$u2.Status; } else { "-"; };

                $pwd1 = if ($u1) {$u1."Letzte Kennwortaenderung"; } else { "-"; };
                $pwd2 = if ($u2) {$u2."Letzte Kennwortaenderung"; } else { "-"; };

                $login1 = if ($u1) {$u1."Neuester Login"; } else { "-"; };
                $login2 = if ($u2) {$u2."Neuester Login"; } else { "-"; };

                $diffType = "Identisch";
                if ($null -eq $u1 -and$null -ne $u2) {$diffType = "Neu erstellt (Nur in Snapshot 2)";
                } elseif ($null -ne $u1 -and$null -eq $u2) {$diffType = "Geloescht (Nur in Snapshot 1)";
                } else {
                    $changes = @();
                    if ($status1 -ne $status2) {$changes += "Status ($status1 ->$status2)"; }
                    if ($pwd1 -ne $pwd2) {$changes += "PW geaendert"; }
                    if ($login1 -ne $login2) {$changes += "Neuer Login"; }

                    if ($changes.Count -gt 0) {
                        $diffType =$changes -join " | ";
                    }
                }

                [void]$diffRows.Add([PSCustomObject]@{
                    "Vergleichsstatus"   = $diffType;
                    "Benutzername"       = $sam;
                    "Anzeigename"        = $disp;
                    "Status (Stand 1)"   = $status1;
                    "Status (Stand 2)"   = $status2;
                    "PW-Datum (Stand 1)" = $pwd1;
                    "PW-Datum (Stand 2)" = $pwd2;
                    "Login (Stand 1)"    = $login1;
                    "Login (Stand 2)"    = $login2;
                    "OU Pfad"            = $ou;
                });
            }

            # Vergleichs-Fenster
            $cmpForm = New-Object System.Windows.Forms.Form;
            $cmpForm.Text = "Snapshot-Vergleich: $([System.IO.Path]::GetFileName($ofd1.FileName))  vs  $([System.IO.Path]::GetFileName($ofd2.FileName))";
            $cmpForm.Size = New-Object System.Drawing.Size(1200, 700);$cmpForm.StartPosition = "CenterParent";
            $cmpForm.Font = New-Object System.Drawing.Font("Segoe UI", 9);

            $pnlCmpTop = New-Object System.Windows.Forms.Panel;
            $pnlCmpTop.Dock = [System.Windows.Forms.DockStyle]::Top;
            $pnlCmpTop.Height = 45;
            $cmpForm.Controls.Add($pnlCmpTop);

            $chkOnlyDiffs = New-Object System.Windows.Forms.CheckBox;
            $chkOnlyDiffs.Text = "Nur geaenderte / neue Konten anzeigen";
            $chkOnlyDiffs.Location = New-Object System.Drawing.Point(15, 12);
            $chkOnlyDiffs.AutoSize =$true;
            $chkOnlyDiffs.Checked =$true;
            $pnlCmpTop.Controls.Add($chkOnlyDiffs);

            $btnExportCmp = New-Object System.Windows.Forms.Button;
            $btnExportCmp.Text = "Vergleich exportieren...";
            $btnExportCmp.Location = New-Object System.Drawing.Point(300, 8);$btnExportCmp.Size = New-Object System.Drawing.Size(160, 28);
            $pnlCmpTop.Controls.Add($btnExportCmp);

            $gridCmp = New-Object System.Windows.Forms.DataGridView;
            $gridCmp.Dock = [System.Windows.Forms.DockStyle]::Fill;
            $gridCmp.ReadOnly =$true;
            $gridCmp.AllowUserToAddRows =$false;
            $gridCmp.RowHeadersVisible =$false;
            $gridCmp.BackgroundColor = [System.Drawing.Color]::White;
            $gridCmp.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells;
            $cmpForm.Controls.Add($gridCmp);$gridCmp.BringToFront();

            $updateCmpGrid = {$filteredRows = [System.Collections.ArrayList]::new();
                if ($chkOnlyDiffs.Checked) {$matched = @($diffRows).Where({$_.Vergleichsstatus -ne "Identisch"; });
                    foreach ($m in $matched) { [void]$filteredRows.Add($m); }
                } else {
                    $filteredRows =$diffRows;
                }
                $gridCmp.DataSource =$filteredRows;

                for ($cr = 0; $cr -lt$gridCmp.Rows.Count; $cr++) {$r = $gridCmp.Rows[$cr];
                    $st =$r.Cells["Vergleichsstatus"].Value;
                    if ($st -like "*Neu erstellt*") {
                        $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(220, 252, 231);$r.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen;
                    } elseif ($st -like "*Geloescht*") {
                        $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 226, 226);$r.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed;
                    } elseif ($st -ne "Identisch") {
                        $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 249, 195);$r.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(146, 64, 14);
                    }
                }
            };

            $chkOnlyDiffs.Add_CheckedChanged({ &$updateCmpGrid; });
            & $updateCmpGrid;

            $btnExportCmp.Add_Click({$sfdCmp = New-Object System.Windows.Forms.SaveFileDialog;
                $sfdCmp.Filter = "CSV-Datei (*.csv)|*.csv";
                $sfdCmp.FileName = "Snapshot_Vergleich_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv";
                if ($sfdCmp.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $cmpData = @($gridCmp.DataSource);
                    $cmpData | Export-Csv -Path$sfdCmp.FileName -Delimiter ";" -NoTypeInformation -Encoding UTF8;
                    [System.Windows.Forms.MessageBox]::Show("Vergleich exportiert!", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information);
                }
            });

            [void]$cmpForm.ShowDialog();

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Vergleich: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
        }
    });

    # SORTIERUNG PER KLICK
    $gridResults.Add_ColumnHeaderMouseClick({
        param($sender, $e);$col = $gridResults.Columns[$e.ColumnIndex];
        if (-not $col) { return; }
        $propName = if ($col.DataPropertyName) { $col.DataPropertyName; } else {$col.HeaderText; };

        if (-not $gridResults.Tag -or -not ($gridResults.Tag -is [hashtable])) {
            $gridResults.Tag = @{ LastCol = ""; Asc = $true };
        }
        $state =$gridResults.Tag;
        if ($state.LastCol -eq$propName) {
            $state.Asc = -not$state.Asc;
        } else {
            $state.LastCol =$propName;
            $state.Asc =$true;
        }

        if ($gridResults.DataSource) {$currentData = @($gridResults.DataSource);$sortObj = if ($state.Asc) {$currentData | Sort-Object { $_.$propName };
            } else {
                $currentData | Sort-Object { $_.$propName } -Descending;
            };
            $arr = [System.Collections.ArrayList]::new();
            foreach ($item in $sortObj) { [void]$arr.Add($item); }
            $gridResults.DataSource =$arr;
            & $applyRowColors;
        }
    });

    $cmbStatusFilter.Add_SelectedIndexChanged({ &$applyFilterAndSort; });
    $cmbSort.Add_SelectedIndexChanged({ &$applyFilterAndSort; });

    $txtSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress =$true;
            $btnRun.PerformClick();
        }
    });

    [void]$subForm.ShowDialog();
}