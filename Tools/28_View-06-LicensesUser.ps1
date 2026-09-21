Add-Type -AssemblyName System.Windows.Forms;
Add-Type -AssemblyName System.Drawing;
[System.Windows.Forms.Application]::EnableVisualStyles();

# Farbpalette
$colNavy    = [System.Drawing.Color]::FromArgb(15, 23, 42);
$colBlue    = [System.Drawing.Color]::FromArgb(0, 120, 212);
$colGreen   = [System.Drawing.Color]::FromArgb(16, 185, 129);
$colWarning = [System.Drawing.Color]::FromArgb(245, 158, 11);
$colDanger  = [System.Drawing.Color]::FromArgb(225, 29, 72);
$colBg      = [System.Drawing.Color]::FromArgb(241, 245, 249);

# Pfadbestimmung fuer TenantData
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { "C:\install\AdminSuite\Tools" };
$c1 = Join-Path -Path $scriptDir -ChildPath "TenantData";
$c2 = "C:\install\AdminSuite\Tools\TenantData";
$c3 = Join-Path -Path (Get-Location).Path -ChildPath "TenantData";
$candidates = @($c1, $c2, $c3);
$tenantBase = $null;
for ($i = 0; $i -lt $candidates.Count; $i++) {
    if (Test-Path $candidates[$i]) {
        $tenantBase = $candidates[$i];
        break;
    }
}
if (-not $tenantBase) { $tenantBase = "C:\install\AdminSuite\Tools\TenantData"; };

# Hauptfenster
$Form = New-Object System.Windows.Forms.Form;
$Form.Text = "28: M365 Lizenz-Matrix & Doppelzuweisungs-Audit (06b_LicensesDetails)";
$Form.Size = New-Object System.Drawing.Size(1450, 920);
$Form.MinimumSize = New-Object System.Drawing.Size(1150, 750);$Form.StartPosition = "CenterScreen";
$Form.BackColor =$colBg;
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 9);

# Header Bar
$TopPanel = New-Object System.Windows.Forms.Panel;
$TopPanel.Dock = "Top";
$TopPanel.Height = 82;
$TopPanel.BackColor =$colNavy;
$Form.Controls.Add($TopPanel);

$lblTitle = New-Object System.Windows.Forms.Label;
$lblTitle.Text = "M365 Lizenz-Matrix & Doppelzuweisungs-Audit";
$lblTitle.ForeColor = [System.Drawing.Color]::White;
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold);
$lblTitle.Location = New-Object System.Drawing.Point(15, 10);
$lblTitle.AutoSize =$true;
$TopPanel.Controls.Add($lblTitle);

$lblPathInfo = New-Object System.Windows.Forms.Label;
$lblPathInfo.Text = "Basis: $tenantBase";
$lblPathInfo.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184);$lblPathInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5);
$lblPathInfo.Location = New-Object System.Drawing.Point(16, 35);
$lblPathInfo.AutoSize =$true;
$TopPanel.Controls.Add($lblPathInfo);

# Stand A Dropdown & Manuell
$lblA = New-Object System.Windows.Forms.Label;
$lblA.Text = "Stand A:";
$lblA.ForeColor = [System.Drawing.Color]::White;
$lblA.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold);
$lblA.Location = New-Object System.Drawing.Point(15, 54);
$lblA.AutoSize =$true;
$TopPanel.Controls.Add($lblA);

$cmbSnapA = New-Object System.Windows.Forms.ComboBox;
$cmbSnapA.DropDownStyle = "DropDownList";
$cmbSnapA.Location = New-Object System.Drawing.Point(75, 51);$cmbSnapA.Size = New-Object System.Drawing.Size(320, 26);
$TopPanel.Controls.Add($cmbSnapA);

$btnBrowseA = New-Object System.Windows.Forms.Button;
$btnBrowseA.Text = "CSV...";
$btnBrowseA.Location = New-Object System.Drawing.Point(400, 50);$btnBrowseA.Size = New-Object System.Drawing.Size(65, 27);
$btnBrowseA.BackColor = [System.Drawing.Color]::FromArgb(51, 65, 85);$btnBrowseA.ForeColor = [System.Drawing.Color]::White;
$btnBrowseA.FlatStyle = "Flat";
$TopPanel.Controls.Add($btnBrowseA);

# Stand B Dropdown & Manuell
$lblB = New-Object System.Windows.Forms.Label;
$lblB.Text = "Stand B:";
$lblB.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184);$lblB.Font = New-Object System.Drawing.Font("Segoe UI", 8.5);
$lblB.Location = New-Object System.Drawing.Point(485, 54);
$lblB.AutoSize =$true;
$TopPanel.Controls.Add($lblB);

$cmbSnapB = New-Object System.Windows.Forms.ComboBox;
$cmbSnapB.DropDownStyle = "DropDownList";
$cmbSnapB.Location = New-Object System.Drawing.Point(545, 51);$cmbSnapB.Size = New-Object System.Drawing.Size(320, 26);
$TopPanel.Controls.Add($cmbSnapB);

$btnBrowseB = New-Object System.Windows.Forms.Button;
$btnBrowseB.Text = "CSV...";
$btnBrowseB.Location = New-Object System.Drawing.Point(870, 50);$btnBrowseB.Size = New-Object System.Drawing.Size(65, 27);
$btnBrowseB.BackColor = [System.Drawing.Color]::FromArgb(51, 65, 85);$btnBrowseB.ForeColor = [System.Drawing.Color]::White;
$btnBrowseB.FlatStyle = "Flat";
$TopPanel.Controls.Add($btnBrowseB);

$btnCompare = New-Object System.Windows.Forms.Button;
$btnCompare.Text = "Vergleichen";
$btnCompare.Location = New-Object System.Drawing.Point(945, 49);$btnCompare.Size = New-Object System.Drawing.Size(100, 28);
$btnCompare.BackColor =$colWarning;
$btnCompare.ForeColor = [System.Drawing.Color]::FromArgb(30, 41, 59);$btnCompare.FlatStyle = "Flat";
$btnCompare.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold);
$TopPanel.Controls.Add($btnCompare);

$btnExport = New-Object System.Windows.Forms.Button;
$btnExport.Text = "CSV Export";
$btnExport.Location = New-Object System.Drawing.Point(1300, 48);$btnExport.Size = New-Object System.Drawing.Size(110, 29);
$btnExport.BackColor =$colGreen;
$btnExport.ForeColor = [System.Drawing.Color]::White;
$btnExport.FlatStyle = "Flat";
$btnExport.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold);
$TopPanel.Controls.Add($btnExport);

# Linkes Filter-Panel
$LeftFilterPanel = New-Object System.Windows.Forms.Panel;
$LeftFilterPanel.Dock = "Left";
$LeftFilterPanel.Width = 320;
$LeftFilterPanel.BackColor = [System.Drawing.Color]::White;
$LeftFilterPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle;
$Form.Controls.Add($LeftFilterPanel);

$lblFilterHead = New-Object System.Windows.Forms.Label;
$lblFilterHead.Text = "LIZENZ- & SPALTEN-FILTER";
$lblFilterHead.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold);
$lblFilterHead.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105);$lblFilterHead.Location = New-Object System.Drawing.Point(10, 10);
$lblFilterHead.AutoSize =$true;
$LeftFilterPanel.Controls.Add($lblFilterHead);

$chkIgnoreFree = New-Object System.Windows.Forms.CheckBox;
$chkIgnoreFree.Text = "Kostenlose ausblenden (Free/Flow)";
$chkIgnoreFree.Checked =$true;
$chkIgnoreFree.Location = New-Object System.Drawing.Point(12, 35);
$chkIgnoreFree.Size = New-Object System.Drawing.Size(295, 22);$chkIgnoreFree.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold);
$LeftFilterPanel.Controls.Add($chkIgnoreFree);

$chkOnlyDups = New-Object System.Windows.Forms.CheckBox;
$chkOnlyDups.Text = "Nur Benutzer mit Doppelzuweisung";
$chkOnlyDups.Checked =$false;
$chkOnlyDups.ForeColor = [System.Drawing.Color]::FromArgb(190, 24, 93);$chkOnlyDups.Location = New-Object System.Drawing.Point(12, 58);
$chkOnlyDups.Size = New-Object System.Drawing.Size(295, 22);$chkOnlyDups.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold);
$LeftFilterPanel.Controls.Add($chkOnlyDups);

$btnSelCopilot = New-Object System.Windows.Forms.Button;
$btnSelCopilot.Text = "Nur Copilot";
$btnSelCopilot.Location = New-Object System.Drawing.Point(12, 85);
$btnSelCopilot.Size = New-Object System.Drawing.Size(95, 26);$btnSelCopilot.FlatStyle = "Flat";
$LeftFilterPanel.Controls.Add($btnSelCopilot);

$btnSelAllPaid = New-Object System.Windows.Forms.Button;
$btnSelAllPaid.Text = "Nur Bezahlte";
$btnSelAllPaid.Location = New-Object System.Drawing.Point(112, 85);
$btnSelAllPaid.Size = New-Object System.Drawing.Size(95, 26);$btnSelAllPaid.FlatStyle = "Flat";
$LeftFilterPanel.Controls.Add($btnSelAllPaid);

$btnSelAll = New-Object System.Windows.Forms.Button;
$btnSelAll.Text = "Alle";
$btnSelAll.Location = New-Object System.Drawing.Point(212, 85);
$btnSelAll.Size = New-Object System.Drawing.Size(85, 26);$btnSelAll.FlatStyle = "Flat";
$LeftFilterPanel.Controls.Add($btnSelAll);

$lblLicList = New-Object System.Windows.Forms.Label;
$lblLicList.Text = "Angezeigte Lizenzspalten:";
$lblLicList.Location = New-Object System.Drawing.Point(10, 118);
$lblLicList.AutoSize =$true;
$lblLicList.Font = New-Object System.Drawing.Font("Segoe UI", 8);
$LeftFilterPanel.Controls.Add($lblLicList);

$chkListLicenses = New-Object System.Windows.Forms.CheckedListBox;
$chkListLicenses.Location = New-Object System.Drawing.Point(12, 138);$chkListLicenses.Size = New-Object System.Drawing.Size(292, 650);
$chkListLicenses.CheckOnClick =$true;
$chkListLicenses.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle;
$LeftFilterPanel.Controls.Add($chkListLicenses);

# Tabs
$TabControl = New-Object System.Windows.Forms.TabControl;
$TabControl.Dock = "Fill";
$TabControl.Padding = New-Object System.Drawing.Point(15, 6);$Form.Controls.Add($TabControl);$TabControl.BringToFront();

$TabSummary    = New-Object System.Windows.Forms.TabPage; $TabSummary.Text = " 1. Lizenzbedarf (Bereinigt) ";
$TabUsers      = New-Object System.Windows.Forms.TabPage; $TabUsers.Text = " 2. Benutzer-Matrix (Lizenzspalten) ";
$TabDuplicates = New-Object System.Windows.Forms.TabPage; $TabDuplicates.Text = " 3. Doppelzuweisungen (Direkt + Gruppe) ";
$TabDelta      = New-Object System.Windows.Forms.TabPage; $TabDelta.Text = " 4. Versionsvergleich (Delta) ";
$TabControl.Controls.AddRange(@($TabSummary,$TabUsers, $TabDuplicates,$TabDelta));

function Build-Grid {
    $grid = New-Object System.Windows.Forms.DataGridView;
    $grid.Dock = "Fill";
    $grid.ReadOnly =$true;
    $grid.BackgroundColor = [System.Drawing.Color]::White;
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None;
    $grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal;
    $grid.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::None;
    $grid.RowHeadersVisible =$false;
    $grid.EnableHeadersVisualStyles =$false;
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect;
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill;
    $grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252);
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105);$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold);
    $grid.ColumnHeadersHeight = 32;
    $grid.RowsDefaultCellStyle.BackColor = [System.Drawing.Color]::White;
    $grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252);
    $grid.RowsDefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(224, 242, 254);$grid.RowsDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(12, 74, 96);
    return $grid;
}

$gridSummary    = Build-Grid; $TabSummary.Controls.Add($gridSummary);$gridUsers      = Build-Grid; $TabUsers.Controls.Add($gridUsers);
$gridDuplicates = Build-Grid; $TabDuplicates.Controls.Add($gridDuplicates);$gridDelta      = Build-Grid; $TabDelta.Controls.Add($gridDelta);

function Bind-DataTable($grid, $items) {$dt = New-Object System.Data.DataTable;
    $arr = @($items);
    if ($arr.Count -gt 0) {
        $first =$arr[0];
        $propNames = @($first.PSObject.Properties.Name);
        for ($p = 0; $p -lt $propNames.Count; $p++) {
            [void]$dt.Columns.Add([string]$propNames[$p]);
        }
        for ($r = 0; $r -lt$arr.Count; $r++) {$dr = $dt.NewRow();$rowItem = $arr[$r];
            for ($c = 0; $c -lt$propNames.Count; $c++) {$col = [string]$propNames[$c];
                $dr[$col] = [string]$rowItem.$col;
            }
            [void]$dt.Rows.Add($dr);
        }
    }
    $grid.DataSource =$dt;
}

function Import-SmartCsv($filePath) {
    if (-not (Test-Path $filePath)) { return$null; }
    $firstLine = Get-Content -Path$filePath -TotalCount 1 -Encoding UTF8;
    $delim = ",";
    if ($firstLine -match ";") { $delim = ";"; }
    $imported = Import-Csv -Path $filePath -Delimiter$delim -Encoding UTF8;
    return @($imported);
}

function Process-RawData($filePath) {
    if (-not $filePath -or -not (Test-Path $filePath)) { return$null; }
    $rawAll = Import-SmartCsv$filePath;
    if (-not $rawAll -or $rawAll.Count -eq 0) { return$null; }

    $clean = [System.Collections.Generic.List[PSCustomObject]]::new();
    for ($idx = 0; $idx -lt$rawAll.Count; $idx++) {$r = $rawAll[$idx];
        $upn = [string]$r.UserPrincipalName;
        if ([string]::IsNullOrWhiteSpace($upn)) { continue; }

        $pName = [string]$r.Produktname;
        $skuPart = [string]$r.SKU_PartNumber;
        if ($skuPart -eq "SPE_F1" -or $pName -eq "Microsoft 365 F1") { $pName = "Microsoft 365 F3"; }
        if ([string]::IsNullOrWhiteSpace($pName)) {$pName = if ($skuPart) {$skuPart } else { "Unbekannt"; }; }
        $isFree = ($skuPart -in @("FLOW_FREE", "POWER_BI_STANDARD") -or $pName -match "Free");

        $clean.Add([PSCustomObject]@{
            UserPrincipalName = $upn.Trim();
            DisplayName       = if ($r.DisplayName) { [string]$r.DisplayName } else {$upn };
            Produktname       = $pName;
            SKU_PartNumber    = $skuPart;
            Zuweisungsart     = if ($r.Zuweisungsart) { [string]$r.Zuweisungsart } else { "Direkt"; };
            Lizenzgruppe_Name = if ($r.Lizenzgruppe_Name) { [string]$r.Lizenzgruppe_Name } else { "-"; };
            SkuId             = if ($r.SkuId) { [string]$r.SkuId } else {$skuPart };
            IsFree            = $isFree;
        });
    }
    return $clean;
}

$global:IsUpdatingFilter =$false;
function Apply-FilterAndRender {
    if (-not $global:CurrentRawRows -or$global:IsUpdatingFilter) { return; }

    $selectedLicenses = [System.Collections.Generic.List[string]]::new();
    for ($i = 0; $i -lt $chkListLicenses.CheckedItems.Count; $i++) {
        [void]$selectedLicenses.Add([string]$chkListLicenses.CheckedItems[$i]);
    }

    $ignoreFree = $chkIgnoreFree.Checked;
    $onlyDups   =$chkOnlyDups.Checked;

    # 1. Doppelzuweisungen
    $userSkuMap = @{};
    for ($i = 0; $i -lt $global:CurrentRawRows.Count; $i++) {
        $item =$global:CurrentRawRows[$i];$key = "$($item.UserPrincipalName)|$($item.SkuId)";
        if (-not $userSkuMap.ContainsKey($key)) {
            $userSkuMap[$key] = [System.Collections.Generic.List[PSCustomObject]]::new();
        }
        [void]$userSkuMap[$key].Add($item);
    }

    $dupsList = [System.Collections.Generic.List[PSCustomObject]]::new();$dupUserMap = @{};
    $mapKeyList = @($userSkuMap.Keys);

    for ($k = 0; $k -lt $mapKeyList.Count; $k++) {
        $key = [string]$mapKeyList[$k];$items = $userSkuMap[$key];
        $hasDirect =$false;
        $hasGroup =$false;
        $groupNames = [System.Collections.Generic.List[string]]::new();

        for ($m = 0; $m -lt$items.Count; $m++) {$cur = $items[$m];
            if ($cur.Zuweisungsart -like "*Direkt*") { $hasDirect =$true; }
            if ($cur.Zuweisungsart -like "*Gruppe*") {
                $hasGroup =$true;
                if ($cur.Lizenzgruppe_Name -and$cur.Lizenzgruppe_Name -ne "-") {
                    [void]$groupNames.Add($cur.Lizenzgruppe_Name);
                }
            }
        }

        if ($hasDirect -and $hasGroup) {$first = $items[0];$grpStr = if ($groupNames.Count -gt 0) { ($groupNames -join ", ") } else { "Gruppenbasiert"; };
            $dupsList.Add([PSCustomObject]@{
                "DisplayName"       = $first.DisplayName;
                "UserPrincipalName" = $first.UserPrincipalName;
                "Produktname"       = $first.Produktname;
                "SKU_PartNumber"    = $first.SKU_PartNumber;
                "Lizenzgruppe_Name" = $grpStr;
                "Status"            = "Doppelt zugewiesen (Direkt + Gruppe)";
            });
            $dupUserMap[$first.UserPrincipalName] =$true;
        }
    }
    $dupsList = @($dupsList | Sort-Object DisplayName);
    Bind-DataTable $gridDuplicates$dupsList;

    # 2. Bereinigter Bedarf
    $skuGroupMap = @{};
    $skuRawCount = @{};$skuUserSet  = @{};

    for ($i = 0; $i -lt$global:CurrentRawRows.Count; $i++) {$r = $global:CurrentRawRows[$i];
        if ($ignoreFree -and $r.IsFree) { continue; }
        if ($selectedLicenses.Count -gt 0 -and ($selectedLicenses -notcontains$r.Produktname)) { continue; }

        $sId = [string]$r.SkuId;
        if (-not $skuGroupMap.ContainsKey($sId)) {$skuGroupMap[$sId] =$r;
            $skuRawCount[$sId] = 0;
            $skuUserSet[$sId]  = @{};
        }
        $skuRawCount[$sId]++;
        $skuUserSet[$sId][$r.UserPrincipalName] =$true;
    }

    $summaryList = [System.Collections.Generic.List[PSCustomObject]]::new();
    $skuKeys = @($skuGroupMap.Keys);
    for ($s = 0; $s -lt $skuKeys.Count; $s++) {
        $sId = [string]$skuKeys[$s];$first = $skuGroupMap[$sId];
        $rawCnt = [int]$skuRawCount[$sId];$netCnt = [int]$skuUserSet[$sId].Keys.Count;
        $dupCnt = $rawCnt -$netCnt;

        $summaryList.Add([PSCustomObject]@{
            "Produktname"           = $first.Produktname;
            "SKU_PartNumber"        = $first.SKU_PartNumber;
            "Bedarf_Bereinigt"      = $netCnt;
            "Roh_Zuweisungen"       = $rawCnt;
            "Doppelt_Direkt_Gruppe" = $dupCnt;
            "Kostenlos"             = if ($first.IsFree) { "Ja" } else { "Nein" };
            "SkuId"                 = $first.SkuId;
        });
    }
    $summaryList = @($summaryList | Sort-Object Bedarf_Bereinigt -Descending);
    Bind-DataTable $gridSummary$summaryList;

    # 3. Benutzer-Matrix (Lizenzspalten)
    $userMap = @{};
    for ($i = 0; $i -lt$global:CurrentRawRows.Count; $i++) {$r = $global:CurrentRawRows[$i];
        $u = [string]$r.UserPrincipalName;
        if (-not $userMap.ContainsKey($u)) {
            $userMap[$u] = [System.Collections.Generic.List[PSCustomObject]]::new();
        }
        [void]$userMap[$u].Add($r);
    }

    $activeLicColumns = [System.Collections.Generic.List[string]]::new();
    if ($selectedLicenses.Count -gt 0) {
        for ($l = 0; $l -lt$selectedLicenses.Count; $l++) {$activeLicColumns.Add($selectedLicenses[$l]);
        }
    } else {
        $distinctP = @{};
        for ($i = 0; $i -lt$global:CurrentRawRows.Count; $i++) {$r = $global:CurrentRawRows[$i];
            if (-not ($ignoreFree -and $r.IsFree)) {$distinctP[$r.Produktname] =$true;
            }
        }
        $pSorted = @($distinctP.Keys | Sort-Object);
        for ($p = 0; $p -lt$pSorted.Count; $p++) {$activeLicColumns.Add([string]$pSorted[$p]);
        }
    }

    $matrixList = [System.Collections.Generic.List[PSCustomObject]]::new();
    $userKeys = @($userMap.Keys);
    for ($u = 0; $u -lt$userKeys.Count; $u++) {$upnKey = [string]$userKeys[$u];
        $uRows =$userMap[$upnKey];$first = $uRows[0];$isUserDup = $dupUserMap.ContainsKey($upnKey);

        if ($onlyDups -and -not$isUserDup) { continue; }

        $rowObj = [ordered]@{
            "DisplayName"     = $first.DisplayName;
            "Doppelzuweisung" = if ($isUserDup) { "JA (Direkt + Gruppe)" } else { "Nein" };
        };

        $hasAnySelected =$false;
        for ($c = 0; $c -lt$activeLicColumns.Count; $c++) {$licName = [string]$activeLicColumns[$c];
            $hasDir =$false;
            $hasGrp =$false;
            $foundAny =$false;

            for ($r = 0; $r -lt$uRows.Count; $r++) {$rowItem = $uRows[$r];
                if ($rowItem.Produktname -eq$licName) {
                    $foundAny =$true;
                    if ($rowItem.Zuweisungsart -like "*Direkt*") { $hasDir =$true; }
                    if ($rowItem.Zuweisungsart -like "*Gruppe*") { $hasGrp =$true; }
                }
            }

            if (-not $foundAny) {
                $rowObj[$licName] = "-";
            } else {
                $hasAnySelected =$true;
                if ($hasDir -and$hasGrp) {
                    $rowObj[$licName] = "Direkt + Gruppe";
                } elseif ($hasGrp) {
                    $rowObj[$licName] = "Gruppenbasiert";
                } else {
                    $rowObj[$licName] = "Direkt";
                }
            }
        }

        if ($hasAnySelected -or$onlyDups) {
            $matrixList.Add([PSCustomObject]$rowObj);
        }
    }
    $matrixList = @($matrixList | Sort-Object DisplayName);
    Bind-DataTable $gridUsers$matrixList;

    $global:LastSummary =$summaryList;
    $global:LastUsers   =$matrixList;
    $global:LastDups    =$dupsList;
}

# Filter-Events
$chkIgnoreFree.Add_CheckedChanged({ Apply-FilterAndRender; });
$chkOnlyDups.Add_CheckedChanged({ Apply-FilterAndRender; });$chkListLicenses.Add_SelectedIndexChanged({
    if ($Form.IsHandleCreated) {$Form.BeginInvoke([Action]{ Apply-FilterAndRender; });
    } else {
        Apply-FilterAndRender;
    }
});

$btnSelCopilot.Add_Click({
    $global:IsUpdatingFilter =$true;
    for ($i = 0; $i -lt $chkListLicenses.Items.Count; $i++) {
        $name = [string]$chkListLicenses.Items[$i];$chkListLicenses.SetItemChecked($i, ($name -match "Copilot"));
    }
    $global:IsUpdatingFilter =$false;
    Apply-FilterAndRender;
});

$btnSelAllPaid.Add_Click({
    $global:IsUpdatingFilter =$true;
    for ($i = 0; $i -lt$chkListLicenses.Items.Count; $i++) {$name = [string]$chkListLicenses.Items[$i];
        $isFree = ($name -match "Free" -or $name -match "FLOW_FREE" -or $name -match "POWER_BI_STANDARD");
        $chkListLicenses.SetItemChecked($i, (-not$isFree));
    }
    $global:IsUpdatingFilter =$false;
    Apply-FilterAndRender;
});

$btnSelAll.Add_Click({
    $global:IsUpdatingFilter =$true;
    for ($i = 0; $i -lt$chkListLicenses.Items.Count; $i++) {$chkListLicenses.SetItemChecked($i,$true);
    }
    $global:IsUpdatingFilter =$false;
    Apply-FilterAndRender;
});

# Datei-Erkennung (matcht 06b_LicensesDetails, 06b_LicenseDetail etc.)
$global:FileMap = @{};
function Scan-TenantFiles {
    $cmbSnapA.Items.Clear();
    $cmbSnapB.Items.Clear();$global:FileMap.Clear();

    if (-not (Test-Path $tenantBase)) {$fbd = New-Object System.Windows.Forms.FolderBrowserDialog;
        $fbd.Description = "Ordner TenantData auswaehlen";
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:tenantBase =$fbd.SelectedPath;
            $lblPathInfo.Text = "Basis: $tenantBase";
        }
    }

    if (Test-Path $tenantBase) {
        $files = @(Get-ChildItem -Path$tenantBase -Filter "*06*License*Detail*.csv" -Recurse | Sort-Object LastWriteTime -Descending);
        for ($f = 0; $f -lt $files.Count; $f++) {
            $fileItem =$files[$f];$rel = $fileItem.FullName.Replace($tenantBase, "").TrimStart("\/");
            $label = "$rel ($($fileItem.LastWriteTime.ToString('yyyy-MM-dd HH:mm')))";
            $global:FileMap[$label] =$fileItem.FullName;
            [void]$cmbSnapA.Items.Add($label);
            [void]$cmbSnapB.Items.Add($label);
        }
    }

    if ($cmbSnapA.Items.Count -gt 0) {$cmbSnapA.SelectedIndex = 0;
        if ($cmbSnapB.Items.Count -gt 1) { $cmbSnapB.SelectedIndex = 1; } else {$cmbSnapB.SelectedIndex = 0; }
    }
}

function Load-ActiveFile($path) {
    if (-not $path -or -not (Test-Path$path)) { return; }
    $global:CurrentRawRows = Process-RawData$path;
    if (-not $global:CurrentRawRows) { return; }

    $global:IsUpdatingFilter =$true;
    $chkListLicenses.Items.Clear();$distinctP = @{};
    for ($i = 0; $i -lt $global:CurrentRawRows.Count; $i++) {
        $distinctP[$global:CurrentRawRows[$i].Produktname] =$true;
    }
    $allProducts = @($distinctP.Keys | Sort-Object);

    for ($p = 0; $p -lt$allProducts.Count; $p++) {$prodName = [string]$allProducts[$p];
        $isFree = ($prodName -match "Free" -or $prodName -match "FLOW_FREE" -or $prodName -match "POWER_BI_STANDARD");
        $defaultCheck = if ($chkIgnoreFree.Checked -and$isFree) { $false; } else {$true; };
        [void]$chkListLicenses.Items.Add($prodName,$defaultCheck);
    }
    $global:IsUpdatingFilter =$false;

    Apply-FilterAndRender;
    $lblPathInfo.Text = "Aktiv: $path";
}

$cmbSnapA.Add_SelectedIndexChanged({
    $key =$cmbSnapA.SelectedItem;
    if ($key -and $global:FileMap.ContainsKey($key)) {
        Load-ActiveFile $global:FileMap[$key];
    }
});

$btnBrowseA.Add_Click({$ofd = New-Object System.Windows.Forms.OpenFileDialog;
    $ofd.Filter = "CSV Dateien (*06*License*.csv;*.csv)|*06*License*.csv;*.csv";
    if (Test-Path $tenantBase) { $ofd.InitialDirectory =$tenantBase; }
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {$k = "Manuell A: " + [System.IO.Path]::GetFileName($ofd.FileName);$global:FileMap[$k] =$ofd.FileName;
        [void]$cmbSnapA.Items.Insert(0, $k);$cmbSnapA.SelectedIndex = 0;
    }
});

$btnBrowseB.Add_Click({$ofd = New-Object System.Windows.Forms.OpenFileDialog;
    $ofd.Filter = "CSV Dateien (*06*License*.csv;*.csv)|*06*License*.csv;*.csv";
    if (Test-Path $tenantBase) { $ofd.InitialDirectory =$tenantBase; }
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {$k = "Manuell B: " + [System.IO.Path]::GetFileName($ofd.FileName);$global:FileMap[$k] =$ofd.FileName;
        [void]$cmbSnapB.Items.Insert(0, $k);$cmbSnapB.SelectedIndex = 0;
    }
});

# Versionsvergleich (Delta)
$btnCompare.Add_Click({
    $snapA =$cmbSnapA.SelectedItem;
    $snapB =$cmbSnapB.SelectedItem;
    if (-not $snapA -or -not$snapB -or ($snapA -eq$snapB)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte zwei unterschiedliche Staende fuer den Vergleich waehlen.", "Hinweis", 0, 48);
        return;
    }
    $fileA =$global:FileMap[$snapA];$fileB = $global:FileMap[$snapB];
    if (-not (Test-Path $fileA) -or -not (Test-Path$fileB)) { return; }

    $rowsA = Process-RawData$fileA;
    $rowsB = Process-RawData$fileB;

    $skuCountA = @{};
    $skuNameA  = @{};$skuSetA   = @{};
    for ($i = 0; $i -lt$rowsA.Count; $i++) {$r = $rowsA[$i];
        $s = [string]$r.SkuId;
        if (-not $skuSetA.ContainsKey($s)) {$skuSetA[$s] = @{};$skuNameA[$s] =$r.Produktname; }
        $skuSetA[$s][$r.UserPrincipalName] =$true;
    }
    $keysA = @($skuSetA.Keys);
    for ($k = 0; $k -lt$keysA.Count; $k++) {$s = [string]$keysA[$k];
        $skuCountA[$s] = $skuSetA[$s].Keys.Count;
    }

    $skuCountB = @{};
    $skuNameB  = @{};$skuSetB   = @{};
    for ($i = 0; $i -lt$rowsB.Count; $i++) {$r = $rowsB[$i];
        $s = [string]$r.SkuId;
        if (-not $skuSetB.ContainsKey($s)) {$skuSetB[$s] = @{};$skuNameB[$s] =$r.Produktname; }
        $skuSetB[$s][$r.UserPrincipalName] =$true;
    }
    $keysB = @($skuSetB.Keys);
    for ($k = 0; $k -lt$keysB.Count; $k++) {$s = [string]$keysB[$k];
        $skuCountB[$s] = $skuSetB[$s].Keys.Count;
    }

    $allSkusMap = @{};
    for ($k = 0; $k -lt $keysA.Count; $k++) { $allSkusMap[[string]$keysA[$k]] =$true; }
    for ($k = 0; $k -lt $keysB.Count; $k++) { $allSkusMap[[string]$keysB[$k]] =$true; }
    $allSkus = @($allSkusMap.Keys | Sort-Object);

    $deltaList = [System.Collections.Generic.List[PSCustomObject]]::new();
    for ($idx = 0; $idx -lt $allSkus.Count; $idx++) {
        $sku = [string]$allSkus[$idx];$cntA = if ($skuCountA.ContainsKey($sku)) { [int]$skuCountA[$sku]; } else { 0; };
        $cntB = if ($skuCountB.ContainsKey($sku)) { [int]$skuCountB[$sku]; } else { 0; };$diff = $cntA -$cntB;
        $pName = if ($skuNameA.ContainsKey($sku)) { $skuNameA[$sku]; } else { $skuNameB[$sku]; };

        $deltaList.Add([PSCustomObject]@{
            "Produktname"    = $pName;
            "Bedarf Stand A" = $cntA;
            "Bedarf Stand B" = $cntB;
            "Delta (A - B)"  = if ($diff -gt 0) { "+$diff"; } elseif ($diff -lt 0) { "$diff"; } else { "0"; };
            "SkuId"          = $sku;
        });
    }

    Bind-DataTable $gridDelta$deltaList;
    $TabControl.SelectedTab =$TabDelta;
});

# CSV Export
$btnExport.Add_Click({
    if (-not $global:LastSummary) { return; }$fbd = New-Object System.Windows.Forms.FolderBrowserDialog;
    $fbd.Description = "Zielordner fuer Lizenz-Export waehlen";
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {$ts = Get-Date -Format "yyyyMMdd_HHmm";
        $f1 = Join-Path -Path $fbd.SelectedPath -ChildPath "01_Lizenzbedarf_Bereinigt_$ts.csv";
        $f2 = Join-Path -Path $fbd.SelectedPath -ChildPath "02_Benutzer_Matrix_$ts.csv";
        $f3 = Join-Path -Path $fbd.SelectedPath -ChildPath "03_Doppelzuweisungen_$ts.csv";

        $global:LastSummary \vert{} Export-Csv -Path$f1 -NoTypeInformation -Encoding UTF8 -Delimiter ";";
        $global:LastUsers   \vert{} Export-Csv -Path$f2 -NoTypeInformation -Encoding UTF8 -Delimiter ";";
        $global:LastDups    \vert{} Export-Csv -Path$f3 -NoTypeInformation -Encoding UTF8 -Delimiter ";";

        [System.Windows.Forms.MessageBox]::Show("Export erfolgreich abgeschlossen!`n`n$f1`n$f2`n$f3", "Export OK", 0, 64);
    }
});

Scan-TenantFiles;
[void]$Form.ShowDialog();$Form.Dispose();