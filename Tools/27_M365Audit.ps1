<#
.SYNOPSIS
    Tool27_M365Audit.ps1 - M365 Multi-Tenant Migration & Inventory Dashboard
    Reiter: 
      - Voraussetzungen & Module (Prüfung & Installation)
      - Mandanten-Dashboard (KPIs & Geräte)
      - Lizenzen (Subscribed SKUs)
      - Mailboxen (Exchange Online)
      - SharePoint Sites
      - Microsoft Teams (Graph-basiert)
    Inklusive Live-Log-Konsole und CSV-Multi-Export.
#>

Add-Type -AssemblyName System.Windows.Forms;
Add-Type -AssemblyName System.Drawing;
[System.Windows.Forms.Application]::EnableVisualStyles();

# -------------------------------------------------------------------------
# UI Farben & Design
# -------------------------------------------------------------------------
$colPrimary  = [System.Drawing.Color]::FromArgb(15, 23, 42);
$colAccent   = [System.Drawing.Color]::FromArgb(0, 120, 212);$colSuccess  = [System.Drawing.Color]::FromArgb(16, 185, 129);
$colBg       = [System.Drawing.Color]::FromArgb(241, 245, 249);$colCardBg   = [System.Drawing.Color]::White;
$colTextDark = [System.Drawing.Color]::FromArgb(30, 41, 59);$colMuted    = [System.Drawing.Color]::FromArgb(100, 116, 139);

$Form = New-Object System.Windows.Forms.Form;
$Form.Text = "Tool27: M365 Tenant Analyzer & Inventory Dashboard";
$Form.Size = New-Object System.Drawing.Size(1200, 880);
$Form.MinimumSize = New-Object System.Drawing.Size(1050, 750);$Form.StartPosition = "CenterScreen";
$Form.BackColor =$colBg;
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 9);

# Globale Variablen
$global:RequiredModules = @("Microsoft.Graph", "ExchangeOnlineManagement");
$global:ReportLicenses  = @();$global:ReportMailboxes = @();
$global:ReportSites     = @();$global:ReportTeams     = @();

# -------------------------------------------------------------------------
# UI Layout: Header & Footer
# -------------------------------------------------------------------------
$TopPanel = New-Object System.Windows.Forms.Panel;
$TopPanel.Dock = "Top";
$TopPanel.Height = 75;
$TopPanel.BackColor =$colPrimary;
$Form.Controls.Add($TopPanel);

$lblTitle = New-Object System.Windows.Forms.Label;
$lblTitle.Text = "M365 Tenant Inventory & Audit (Migration Prep)";
$lblTitle.ForeColor = [System.Drawing.Color]::White;
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold);
$lblTitle.Location = New-Object System.Drawing.Point(18, 14);
$lblTitle.AutoSize =$true;
$TopPanel.Controls.Add($lblTitle);

$lblSub = New-Object System.Windows.Forms.Label;
$lblSub.Text = "Bestandsaufnahme: Identitaeten, Lizenzen, Exchange, SharePoint und Teams";
$lblSub.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184);$lblSub.Font = New-Object System.Drawing.Font("Segoe UI", 8.5);
$lblSub.Location = New-Object System.Drawing.Point(19, 40);
$lblSub.AutoSize =$true;
$TopPanel.Controls.Add($lblSub);

$lblUser = New-Object System.Windows.Forms.Label;
$lblUser.Text = "Admin UPN:";
$lblUser.ForeColor = [System.Drawing.Color]::White;
$lblUser.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold);
$lblUser.Location = New-Object System.Drawing.Point(540, 26);
$lblUser.AutoSize =$true;
$TopPanel.Controls.Add($lblUser);

$txtUser = New-Object System.Windows.Forms.TextBox;
$txtUser.Location = New-Object System.Drawing.Point(625, 23);
$txtUser.Size = New-Object System.Drawing.Size(250, 27);$txtUser.Font = New-Object System.Drawing.Font("Segoe UI", 9.5);
$TopPanel.Controls.Add($txtUser);

$btnConnect = New-Object System.Windows.Forms.Button;
$btnConnect.Text = "Verbinden & Laden";
$btnConnect.Location = New-Object System.Drawing.Point(885, 21);$btnConnect.Size = New-Object System.Drawing.Size(140, 31);
$btnConnect.BackColor =$colAccent;
$btnConnect.ForeColor = [System.Drawing.Color]::White;
$btnConnect.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat;
$btnConnect.FlatAppearance.BorderSize = 0;
$btnConnect.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold);
$btnConnect.Cursor = [System.Windows.Forms.Cursors]::Hand;
$TopPanel.Controls.Add($btnConnect);

$btnExport = New-Object System.Windows.Forms.Button;
$btnExport.Text = "CSV Export";
$btnExport.Location = New-Object System.Drawing.Point(1035, 21);$btnExport.Size = New-Object System.Drawing.Size(110, 31);
$btnExport.BackColor =$colSuccess;
$btnExport.ForeColor = [System.Drawing.Color]::White;
$btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat;
$btnExport.FlatAppearance.BorderSize = 0;
$btnExport.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold);
$btnExport.Cursor = [System.Windows.Forms.Cursors]::Hand;
$btnExport.Enabled =$false;
$TopPanel.Controls.Add($btnExport);

# Live-Log Konsole
$LogPanel = New-Object System.Windows.Forms.Panel;
$LogPanel.Dock = "Bottom";
$LogPanel.Height = 150;
$LogPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 41, 59);
$Form.Controls.Add($LogPanel);

$lblLogHeader = New-Object System.Windows.Forms.Label;
$lblLogHeader.Text = "AKTIVITAETS- & ABFRAGE-LOG:";
$lblLogHeader.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184);$lblLogHeader.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold);
$lblLogHeader.Location = New-Object System.Drawing.Point(12, 6);
$lblLogHeader.AutoSize =$true;
$LogPanel.Controls.Add($lblLogHeader);

$txtLog = New-Object System.Windows.Forms.TextBox;
$txtLog.Multiline =$true;
$txtLog.ReadOnly =$true;
$txtLog.ScrollBars = "Vertical";
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(15, 23, 42);
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(226, 232, 240);$txtLog.Font = New-Object System.Drawing.Font("Consolas", 8.5);
$txtLog.Dock = "Bottom";
$txtLog.Height = 120;
$LogPanel.Controls.Add($txtLog);

function Write-GuiLog($msg,$level="INFO") {
    $time = Get-Date -Format "HH:mm:ss";
    $txtLog.AppendText("[$time] [$level]$msg`r`n");
    $txtLog.SelectionStart =$txtLog.TextLength;
    $txtLog.ScrollToCaret();$Form.Update();
}

# TabControl
$TabControl = New-Object System.Windows.Forms.TabControl;
$TabControl.Dock = "Fill";
$TabControl.Padding = New-Object System.Drawing.Point(14, 6);$Form.Controls.Add($TabControl);$TabControl.BringToFront();

$TabPrereq = New-Object System.Windows.Forms.TabPage;
$TabPrereq.Text = "  Voraussetzungen & Module  ";
$Tab0      = New-Object System.Windows.Forms.TabPage;
$Tab0.Text      = "  0: Mandanten-Dashboard  ";
$Tab1      = New-Object System.Windows.Forms.TabPage;
$Tab1.Text      = "  1: Lizenzen  ";
$Tab2      = New-Object System.Windows.Forms.TabPage;
$Tab2.Text      = "  2: Mailboxen  ";
$Tab3      = New-Object System.Windows.Forms.TabPage;
$Tab3.Text      = "  3: SharePoint Sites  ";
$Tab4      = New-Object System.Windows.Forms.TabPage;
$Tab4.Text      = "  4: Microsoft Teams  ";

$TabPrereq.BackColor =$colBg;
$Tab0.BackColor      =$colBg;
$Tab1.BackColor      = [System.Drawing.Color]::White;
$Tab2.BackColor      = [System.Drawing.Color]::White;
$Tab3.BackColor      = [System.Drawing.Color]::White;
$Tab4.BackColor      = [System.Drawing.Color]::White;

$TabControl.Controls.AddRange(@($TabPrereq,$Tab0, $Tab1,$Tab2, $Tab3,$Tab4));

# -------------------------------------------------------------------------
# Typensichere Tabellen- und Kachel-Funktionen
# -------------------------------------------------------------------------
function Build-StyledGrid {
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
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105);$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold);
    $grid.ColumnHeadersHeight = 36;
    $grid.RowsDefaultCellStyle.BackColor = [System.Drawing.Color]::White;
    $grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252);
    $grid.RowsDefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(224, 242, 254);$grid.RowsDefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(12, 74, 96);
    return $grid;
}

function Set-GridData($container,$dataSource) {
    $container.Controls.Clear();$grid = Build-StyledGrid;
    $table = New-Object System.Data.DataTable;

    if ($dataSource -and @($dataSource).Count -gt 0) {
        $firstItem = @($dataSource)[0];
        $propList = [System.Collections.ArrayList]@($firstItem.PSObject.Properties.Name);

        for ($c = 0; $c -lt $propList.Count; $c++) {
            [void]$table.Columns.Add([string]$propList[$c]);
        }

        $dataArr = @($dataSource);
        for ($d = 0; $d -lt$dataArr.Count; $d++) {$row = $table.NewRow();$item = $dataArr[$d];
            for ($c = 0; $c -lt $propList.Count; $c++) {
                $pName = [string]$propList[$c];$val = $item.$pName;
                $row[$pName] = if ($null -ne $val) { [string]$val } else { "" };
            }
            [void]$table.Rows.Add($row);
        }
    }
    $grid.DataSource =$table;
    $container.Controls.Add($grid);
}

function New-StatCard($title,$mainValue, $subText,$posX, $posY,$width, $height) {$pnl = New-Object System.Windows.Forms.Panel;
    $pnl.Location = New-Object System.Drawing.Point($posX, $posY);$pnl.Size = New-Object System.Drawing.Size($width,$height);
    $pnl.BackColor =$colCardBg;
    $pnl.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle;

    $lTitle = New-Object System.Windows.Forms.Label;
    $lTitle.Text = $title.ToUpper();$lTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold);
    $lTitle.ForeColor =$colMuted;
    $lTitle.Location = New-Object System.Drawing.Point(12, 10);
    $lTitle.AutoSize =$true;

    $lVal = New-Object System.Windows.Forms.Label;
    $lVal.Text =$mainValue;
    $lVal.Font = New-Object System.Drawing.Font("Segoe UI", 17, [System.Drawing.FontStyle]::Bold);
    $lVal.ForeColor =$colTextDark;
    $lVal.Location = New-Object System.Drawing.Point(10, 26);
    $lVal.AutoSize =$true;

    $lSub = New-Object System.Windows.Forms.Label;
    $lSub.Text =$subText;
    $lSub.Font = New-Object System.Drawing.Font("Segoe UI", 8);
    $lSub.ForeColor =$colMuted;
    $lSub.Location = New-Object System.Drawing.Point(12, 58);
    $lSub.AutoSize =$true;

    $pnl.Controls.AddRange(@($lTitle, $lVal,$lSub));
    return $pnl;
}

# -------------------------------------------------------------------------
# Register: Voraussetzungen & Module
# -------------------------------------------------------------------------
$pnlPrereqTop = New-Object System.Windows.Forms.Panel;
$pnlPrereqTop.Dock = "Top";
$pnlPrereqTop.Height = 160;
$pnlPrereqTop.BackColor =$colCardBg;
$pnlPrereqTop.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle;
$TabPrereq.Controls.Add($pnlPrereqTop);

$lblEnvHeader = New-Object System.Windows.Forms.Label;
$lblEnvHeader.Text = "SYSTEM- & POWERSHELL-UMGEBUNGSPRUEFUNG";
$lblEnvHeader.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold);
$lblEnvHeader.Location = New-Object System.Drawing.Point(15, 12);
$lblEnvHeader.AutoSize =$true;
$pnlPrereqTop.Controls.Add($lblEnvHeader);

$lblPsVer = New-Object System.Windows.Forms.Label;
$lblPsVer.Text = "PowerShell Version: $($PSVersionTable.PSVersion.ToString())";
$lblPsVer.Location = New-Object System.Drawing.Point(18, 42);
$lblPsVer.AutoSize =$true;
$pnlPrereqTop.Controls.Add($lblPsVer);

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator);$lblAdmin = New-Object System.Windows.Forms.Label;
$lblAdmin.Text = if ($isAdmin) { "Administrator-Rechte: JA (Erhoehte Rechte aktiv)" } else { "Administrator-Rechte: NEIN (Standardbenutzer)" };
$lblAdmin.ForeColor = if ($isAdmin) { [System.Drawing.Color]::ForestGreen } else { [System.Drawing.Color]::DarkOrange };$lblAdmin.Location = New-Object System.Drawing.Point(18, 65);
$lblAdmin.AutoSize =$true;
$pnlPrereqTop.Controls.Add($lblAdmin);

$lblExec = New-Object System.Windows.Forms.Label;
$lblExec.Text = "Execution Policy: $(Get-ExecutionPolicy)";
$lblExec.Location = New-Object System.Drawing.Point(18, 88);
$lblExec.AutoSize =$true;
$pnlPrereqTop.Controls.Add($lblExec);

$btnCheckModules = New-Object System.Windows.Forms.Button;
$btnCheckModules.Text = "Pruefung wiederholen";
$btnCheckModules.Location = New-Object System.Drawing.Point(18, 115);$btnCheckModules.Size = New-Object System.Drawing.Size(160, 30);
$btnCheckModules.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249);$btnCheckModules.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat;
$pnlPrereqTop.Controls.Add($btnCheckModules);

$btnInstallMissing = New-Object System.Windows.Forms.Button;
$btnInstallMissing.Text = "Fehlende Module jetzt installieren";
$btnInstallMissing.Location = New-Object System.Drawing.Point(190, 115);$btnInstallMissing.Size = New-Object System.Drawing.Size(230, 30);
$btnInstallMissing.BackColor =$colAccent;
$btnInstallMissing.ForeColor = [System.Drawing.Color]::White;
$btnInstallMissing.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat;
$pnlPrereqTop.Controls.Add($btnInstallMissing);

$pnlModuleGrid = New-Object System.Windows.Forms.Panel;
$pnlModuleGrid.Dock = "Fill";
$pnlModuleGrid.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 0);$TabPrereq.Controls.Add($pnlModuleGrid);$pnlModuleGrid.BringToFront();

function Refresh-ModuleCheck {
    Write-GuiLog "Pruefe Modulstatus fuer: $($global:RequiredModules -join ', ')";
    $modStatusList = @();
    for ($i = 0; $i -lt$global:RequiredModules.Count; $i++) {$name = $global:RequiredModules[$i];
        $found = Get-Module -ListAvailable -Name$name;
        $installed = if ($found) { "Installiert (OK)" } else { "FEHLT" };
        $ver = if ($found) {$found[0].Version.ToString() } else { "-" };
        $modStatusList += [PSCustomObject]@{
            "Modulname"          = $name;
            "Status"             = $installed;
            "Installierte Vers." = $ver;
            "Erforderlich fuer"   = if ($name -like "*Graph*") { "Identitaeten, Lizenzen, SharePoint, Teams" } else { "Mailboxen & Postfachtypen" };
        };
    }
    Set-GridData $pnlModuleGrid$modStatusList;
}

$btnCheckModules.Add_Click({
    Refresh-ModuleCheck;
});

$btnInstallMissing.Add_Click({
    $btnInstallMissing.Enabled =$false;
    for ($i = 0; $i -lt$global:RequiredModules.Count; $i++) {$mName = $global:RequiredModules[$i];
        $chk = Get-Module -ListAvailable -Name$mName;
        if (-not $chk) {
            Write-GuiLog "Starte Installation von $mName..." "WARN";
            try {
                Install-Module -Name $mName -Scope CurrentUser -Force -AllowClobber;
                Write-GuiLog "Modul $mName erfolgreich installiert." "SUCCESS";
            } catch {
                Write-GuiLog "Fehler bei der Installation von $mName :$_" "ERROR";
            }
        }
    }
    Refresh-ModuleCheck;
    $btnInstallMissing.Enabled =$true;
});

# -------------------------------------------------------------------------
# Verbindungs- und Abfragelogik
# -------------------------------------------------------------------------
$btnConnect.Add_Click({
    $targetUser =$txtUser.Text.Trim();
    if ([string]::IsNullOrWhiteSpace($targetUser)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie die Admin-UPN des gewuenschten Mandanten ein!", "Hinweis", 0, 48);
        return;
    }

    $btnConnect.Enabled =$false;
    $btnExport.Enabled  =$false;
    Write-GuiLog "Verbindungsprozess gestartet fuer Admin: $targetUser";

    try {
        # 1. Zentraler Login ueber Microsoft Graph
        Write-GuiLog "Verbinde mit Microsoft Graph (Zentrale Anmeldung)...";
        $graphScopes = @("User.Read.All", "Group.Read.All", "Device.Read.All", "Organization.Read.All", "Sites.Read.All", "Team.ReadBasic.All", "Channel.ReadBasic.All");
        Connect-MgGraph -Scopes $graphScopes -NoWelcome;

        # 2. Exchange Online verbindet sich mit dem Zielkonto
        Write-GuiLog "Verbinde mit Exchange Online...";
        Connect-ExchangeOnline -UserPrincipalName $targetUser -ShowBanner:$false;

        Write-GuiLog "Alle Dienste erfolgreich verbunden. Beginne Mandantenabfrage..." "SUCCESS";

        # Register 0: Dashboard (KPIs & Geraete)
        Write-GuiLog "Lese Benutzer & AD-Sync-Status...";
        $allUsers = Get-MgUser -All -Property Id, DisplayName, OnPremisesSyncEnabled, AccountEnabled;
        $totalUsers   =$allUsers.Count;
        $cloudUsers   = 0;
        $syncedUsers  = 0;
        $enabledUsers = 0;

        for ($u = 0; $u -lt$totalUsers; $u++) {$userObj = $allUsers[$u];
            if ($userObj.OnPremisesSyncEnabled -eq$true) { $syncedUsers++; } else {$cloudUsers++; }
            if ($userObj.AccountEnabled -eq $true) {$enabledUsers++; }
        }

        Write-GuiLog "Lese Gruppen...";
        $allGroups = Get-MgGroup -All -Property Id, DisplayName, OnPremisesSyncEnabled, GroupTypes;
        $totalGroups  =$allGroups.Count;
        $cloudGroups  = 0;
        $syncedGroups = 0;
        for ($g = 0; $g -lt $totalGroups; $g++) {
            if ($allGroups[$g].OnPremisesSyncEnabled -eq$true) { $syncedGroups++; } else {$cloudGroups++; }
        }

        Write-GuiLog "Lese Endgeraete-Inventar...";
        $allDevices = Get-MgDevice -All -Property Id, DisplayName, OperatingSystem, TrustType;
        $totalDevices =$allDevices.Count;
        $osHash = @{};
        for ($dev = 0; $dev -lt$totalDevices; $dev++) {$osName = $allDevices[$dev].OperatingSystem;
            if ([string]::IsNullOrWhiteSpace($osName)) {$osName = "Unbekannt / Sonstige"; }
            if ($osHash.ContainsKey($osName)) {
                $osHash[$osName]++;
            } else {
                $osHash[$osName] = 1;             }         }$deviceSummary = @();
        foreach ($k in$osHash.Keys) {
            $deviceSummary += [PSCustomObject]@{ "Betriebssystem / Kategorie" = $k; "Anzahl Geraete" = $osHash[$k] };
        }

        # UI Register 0
        $Tab0.Controls.Clear();$pnlKpi = New-Object System.Windows.Forms.Panel;
        $pnlKpi.Dock = "Top";
        $pnlKpi.Height = 110;
        $pnlKpi.BackColor =$colBg;
        $Tab0.Controls.Add($pnlKpi);

        $c1 = New-StatCard "Benutzer Gesamt" "$totalUsers" "Aktiv: $enabledUsers" 15 12 240 85;
        $c2 = New-StatCard "Cloud vs. Synced" "$cloudUsers Cloud" "AD-Sync: $syncedUsers" 270 12 240 85;
        $c3 = New-StatCard "Gruppen Gesamt" "$totalGroups" "Cloud: $cloudGroups \vert{} AD:$syncedGroups" 525 12 260 85;
        $c4 = New-StatCard "Geraete-Bestand" "$totalDevices" "Registriert in Entra ID" 800 12 240 85;
        $pnlKpi.Controls.AddRange(@($c1,$c2, $c3,$c4));

        $pnlGridDev = New-Object System.Windows.Forms.Panel;
        $pnlGridDev.Dock = "Fill";
        $pnlGridDev.Padding = New-Object System.Windows.Forms.Padding(15, 0, 15, 15);
        $pnlGridDev.BackColor =$colBg;
        $Tab0.Controls.Add($pnlGridDev);
        $pnlGridDev.BringToFront();$pnlDevInner = New-Object System.Windows.Forms.Panel;
        $pnlDevInner.Dock = "Fill";
        $pnlDevInner.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle;
        $pnlGridDev.Controls.Add($pnlDevInner);
        Set-GridData $pnlDevInner$deviceSummary;

        # Register 1: Lizenzen
        Write-GuiLog "Lese Lizenzkontingente (Subscribed SKUs)...";
        $skus = Get-MgSubscribedSku;
        $global:ReportLicenses = @();
        for ($s = 0; $s -lt $skus.Count; $s++) {
            $sku =$skus[$s];$prep = if ($sku.PrepaidUnits) {$sku.PrepaidUnits.Enabled } else { 0 };
            $cons = if ($sku.ConsumedUnits) { $sku.ConsumedUnits } else { 0 };$global:ReportLicenses += [PSCustomObject]@{
                "SKU PartNumber" = $sku.SkuPartNumber;
                "Gekauft"        = $prep;
                "Zugewiesen"     = $cons;
                "Verfuegbar"     = ($prep -$cons);
                "Gueltig fuer"   = $sku.AppliesTo;
            };
        }
        Set-GridData $Tab1$global:ReportLicenses;

        # Register 2: Mailboxen
        Write-GuiLog "Lese Exchange Online Postfaecher aus...";
        $mbs = Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails, ArchiveStatus;
        $global:ReportMailboxes = @();
        for ($m = 0; $m -lt $mbs.Count; $m++) {
            $mb =$mbs[$m];$global:ReportMailboxes += [PSCustomObject]@{
                "Anzeigename"      = $mb.DisplayName;
                "UPN / SMTP"       = $mb.UserPrincipalName;
                "Postfachtyp"      = $mb.RecipientTypeDetails;
                "Archiv-Status"    = $mb.ArchiveStatus;
                "GAL ausgeblendet" = $mb.HiddenFromAddressListsEnabled;
            };
        }
        Set-GridData $Tab2$global:ReportMailboxes;

        # Register 3: SharePoint Sites
        Write-GuiLog "Lese SharePoint Site Collections aus...";
        $sites = Get-MgSite -All;
        $global:ReportSites = @();
        for ($st = 0; $st -lt $sites.Count; $st++) {
            $site =$sites[$st];$global:ReportSites += [PSCustomObject]@{
                "Site DisplayName" = $site.DisplayName;
                "Site Name"        = $site.Name;
                "Web URL"          = $site.WebUrl;
                "Site ID"          = $site.Id;
            };
        }
        Set-GridData $Tab3$global:ReportSites;

        # Register 4: Microsoft Teams (Reine Graph-API - Ohne WAM-Fehler)
        Write-GuiLog "Lese Microsoft Teams via Graph API aus...";
        $global:ReportTeams = @();
        try {
            $teamsGroups = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -Property Id, DisplayName, Description, Visibility -All -ErrorAction Stop;
            if ($teamsGroups -and @($teamsGroups).Count -gt 0) {
                for ($t = 0; $t -lt @($teamsGroups).Count; $t++) {
                    $tm = @($teamsGroups)[$t];$global:ReportTeams += [PSCustomObject]@{
                        "Team Name"    = [string]$tm.DisplayName;
                        "Sichtbarkeit" = [string]$tm.Visibility;
                        "GroupId"      = [string]$tm.Id;
                        "Beschreibung" = [string]$tm.Description;
                    };
                }
                Write-GuiLog "$(@($teamsGroups).Count) Team(s) gefunden." "INFO";
            } else {
                Write-GuiLog "Keine Microsoft Teams im Mandanten eingerichtet (0 Teams)." "INFO";
                $global:ReportTeams += [PSCustomObject]@{
                    "Hinweis" = "Keine Microsoft Teams im Mandanten gefunden.";
                };
            }
        } catch {
            Write-GuiLog "Warnung bei Teams-Graph-Abfrage: $_" "WARN";
            $global:ReportTeams += [PSCustomObject]@{
                "Fehler" = "Teams konnten nicht abgerufen werden: $_";
            };
        }
        Set-GridData $Tab4$global:ReportTeams;

        Write-GuiLog "Datenabfrage fuer Mandanten vollstaendig abgeschlossen!" "SUCCESS";
        $btnExport.Enabled =$true;
        $TabControl.SelectedTab =$Tab0;
    } catch {
        Write-GuiLog "Fehler bei der Abfrage: $_" "ERROR";
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Abrufen der Tenant-Daten:`n`n$_", "Fehler", 0, 16);
    } finally {
        $btnConnect.Enabled =$true;
    }
});

# -------------------------------------------------------------------------
# CSV Export
# -------------------------------------------------------------------------
$btnExport.Add_Click({$fbd = New-Object System.Windows.Forms.FolderBrowserDialog;
    $fbd.Description = "Ordner fuer Tenant-Export waehlen";
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $p =$fbd.SelectedPath;
        $ts = Get-Date -Format "yyyyMMdd_HHmm";
        Write-GuiLog "Exportiere CSV-Dateien nach: $p";
        
        $global:ReportLicenses  | Export-Csv -Path "$p\M365_Licenses_$ts.csv" -NoTypeInformation -Encoding UTF8;
        $global:ReportMailboxes | Export-Csv -Path "$p\M365_Mailboxes_$ts.csv" -NoTypeInformation -Encoding UTF8;
        $global:ReportSites     | Export-Csv -Path "$p\M365_SharePoint_$ts.csv" -NoTypeInformation -Encoding UTF8;
        $global:ReportTeams     | Export-Csv -Path "$p\M365_Teams_$ts.csv" -NoTypeInformation -Encoding UTF8;
        
        Write-GuiLog "4 CSV-Dateien erfolgreich exportiert." "SUCCESS";
        [System.Windows.Forms.MessageBox]::Show("CSV-Export abgeschlossen!`nDateien liegen in:`n$p", "Export OK", 0, 64);
    }
});

# Initiale Pruefung beim Start
Refresh-ModuleCheck;
Write-GuiLog "M365 Analyzer bereit. Tragen Sie oben die Admin-UPN ein und klicken Sie auf 'Verbinden & Laden'.";

[void]$Form.ShowDialog();$Form.Dispose();
# --- ENDE TOOL27 ---