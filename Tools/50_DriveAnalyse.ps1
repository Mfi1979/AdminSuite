Add-Type -AssemblyName System.Windows.Forms;
Add-Type -AssemblyName System.Drawing;

if (-not [System.Windows.Forms.Application]::RenderWithVisualStyles) {
    try { [System.Windows.Forms.Application]::EnableVisualStyles(); } catch {}
}

# ==============================================================================
# DESIGN-VORGABEN (STANDARD)
# ==============================================================================
$script:UITheme = @{
    HeaderHeight       = 34;
    RowHeight          = 28;
    HeaderPaddingLeft  = 8;
    HeaderPaddingRight = 8;
    CellPaddingLeft    = 8;
    CellPaddingRight   = 8;
    HeaderFont         = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold);
    CellFont           = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Regular);
    HeaderBackColor    = [System.Drawing.Color]::FromArgb(238, 242, 248);
    HeaderForeColor    = [System.Drawing.Color]::FromArgb(30, 41, 59);
    GridLineColor      = [System.Drawing.Color]::FromArgb(226, 232, 240);
    RowBackColor       = [System.Drawing.Color]::White;
    RowAltBackColor    = [System.Drawing.Color]::FromArgb(248, 250, 252);
    SelectionBackColor = [System.Drawing.Color]::FromArgb(203, 228, 249);
    SelectionForeColor = [System.Drawing.Color]::Black;
};

function Apply-StandardGridTheme {
    param([System.Windows.Forms.DataGridView]$grid)
    if (-not $grid) { return; }

    $grid.EnableHeadersVisualStyles = $false;
    $grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing;
    $grid.ColumnHeadersHeight = $script:UITheme.HeaderHeight;
    $grid.RowTemplate.Height = $script:UITheme.RowHeight;

    $hdrStyle = New-Object System.Windows.Forms.DataGridViewCellStyle;
    $hdrStyle.Font = $script:UITheme.HeaderFont;
    $hdrStyle.BackColor = $script:UITheme.HeaderBackColor;
    $hdrStyle.ForeColor = $script:UITheme.HeaderForeColor;
    $hdrStyle.Alignment = [System.Drawing.ContentAlignment]::MiddleLeft;
    $hdrStyle.Padding = New-Object System.Windows.Forms.Padding($script:UITheme.HeaderPaddingLeft, 0, $script:UITheme.HeaderPaddingRight, 0);
    $grid.ColumnHeadersDefaultCellStyle = $hdrStyle;

    $cellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle;
    $cellStyle.Font = $script:UITheme.CellFont;
    $cellStyle.Alignment = [System.Drawing.ContentAlignment]::MiddleLeft;
    $cellStyle.Padding = New-Object System.Windows.Forms.Padding($script:UITheme.CellPaddingLeft, 2, $script:UITheme.CellPaddingRight, 2);
    $cellStyle.SelectionBackColor = $script:UITheme.SelectionBackColor;
    $cellStyle.SelectionForeColor = $script:UITheme.SelectionForeColor;
    $grid.DefaultCellStyle = $cellStyle;

    $grid.AlternatingRowsDefaultCellStyle.BackColor = $script:UITheme.RowAltBackColor;
    $grid.AlternatingRowsDefaultCellStyle.SelectionBackColor = $script:UITheme.SelectionBackColor;
    $grid.AlternatingRowsDefaultCellStyle.SelectionForeColor = $script:UITheme.SelectionForeColor;

    $grid.GridColor = $script:UITheme.GridLineColor;
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None;
    $grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal;
    $grid.BackgroundColor = [System.Drawing.Color]::White;
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect;
    $grid.MultiSelect = $false;
    $grid.AllowUserToAddRows = $false;
    $grid.AllowUserToDeleteRows = $false;
    $grid.ReadOnly = $true;
    $grid.RowHeadersVisible = $false;
}

function Show-StorageAnalyzerGUI {
    [CmdletBinding()]
    param()

    # Interne Datenbehälter
    $script:MasterFolders  = [System.Collections.Generic.List[PSCustomObject]]::new();
    $script:SubFoldersMap  = @{};
    $script:DriveList      = [System.Collections.Generic.List[PSCustomObject]]::new();
    $script:ExpandedRoots  = [System.Collections.Generic.HashSet[string]]::new();
    $script:UserProfiles   = [System.Collections.Generic.List[PSCustomObject]]::new();
    $script:UserFilesMap   = @{};
    $script:RecycleBinData = [System.Collections.Generic.List[PSCustomObject]]::new();
    $script:IsScanning     = $false;

    $form = New-Object System.Windows.Forms.Form;
    $form.Text = "Client Speicherplatz-, Benutzer-, Papierkorb- & Wartungs-Dashboard [v5.0]";
    $form.Size = New-Object System.Drawing.Size(1500, 940);
    $form.StartPosition = "CenterScreen";
    $form.MinimumSize = New-Object System.Drawing.Size(1150, 750);
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9.0);
    $form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250);

    # =========================================================================
    # KOPFZEILE / STEUERUNG
    # =========================================================================
    $tblTop = New-Object System.Windows.Forms.TableLayoutPanel;
    $tblTop.Dock = [System.Windows.Forms.DockStyle]::Top;
    $tblTop.Height = 115;
    $tblTop.BackColor = [System.Drawing.Color]::FromArgb(238, 242, 248);
    $tblTop.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 10);
    $tblTop.ColumnCount = 4;
    $tblTop.RowCount = 1;
    $tblTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 240)));
    $tblTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 260)));
    $tblTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 450)));
    $tblTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)));
    $form.Controls.Add($tblTop);

    # Col 0: Host
    $pnlCol0 = New-Object System.Windows.Forms.Panel;
    $pnlCol0.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $tblTop.Controls.Add($pnlCol0, 0, 0);

    $lblHost = New-Object System.Windows.Forms.Label;
    $lblHost.Text = "Ziel-System:";
    $lblHost.Location = New-Object System.Drawing.Point(0, 4);
    $lblHost.AutoSize = $true;
    $lblHost.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold);
    $pnlCol0.Controls.Add($lblHost);

    $txtHost = New-Object System.Windows.Forms.TextBox;
    $txtHost.Text = "localhost";
    $txtHost.Location = New-Object System.Drawing.Point(90, 2);
    $txtHost.Size = New-Object System.Drawing.Size(135, 25);
    $pnlCol0.Controls.Add($txtHost);

    $btnLoadDrives = New-Object System.Windows.Forms.Button;
    $btnLoadDrives.Text = "Laufwerke laden";
    $btnLoadDrives.Location = New-Object System.Drawing.Point(90, 40);
    $btnLoadDrives.Size = New-Object System.Drawing.Size(135, 34);
    $btnLoadDrives.BackColor = [System.Drawing.Color]::FromArgb(225, 235, 248);
    $btnLoadDrives.FlatStyle = [System.Windows.Forms.FlatStyle]::System;
    $pnlCol0.Controls.Add($btnLoadDrives);

    # Col 1: Laufwerke
    $pnlCol1 = New-Object System.Windows.Forms.Panel;
    $pnlCol1.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $tblTop.Controls.Add($pnlCol1, 1, 0);

    $lblDriveBox = New-Object System.Windows.Forms.Label;
    $lblDriveBox.Text = "Laufwerke:";
    $lblDriveBox.Location = New-Object System.Drawing.Point(5, 4);
    $lblDriveBox.AutoSize = $true;
    $lblDriveBox.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold);
    $pnlCol1.Controls.Add($lblDriveBox);

    $chkListDrives = New-Object System.Windows.Forms.CheckedListBox;
    $chkListDrives.Location = New-Object System.Drawing.Point(85, 2);
    $chkListDrives.Size = New-Object System.Drawing.Size(165, 82);
    $chkListDrives.CheckOnClick = $true;
    $pnlCol1.Controls.Add($chkListDrives);

    # Col 2: Optionen
    $pnlCol2 = New-Object System.Windows.Forms.Panel;
    $pnlCol2.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $tblTop.Controls.Add($pnlCol2, 2, 0);

    $chkFilter1GB = New-Object System.Windows.Forms.CheckBox;
    $chkFilter1GB.Text = "Nur Verzeichnisse >= 1.0 GB anzeigen";
    $chkFilter1GB.Location = New-Object System.Drawing.Point(8, 4);
    $chkFilter1GB.Size = New-Object System.Drawing.Size(320, 22);
    $chkFilter1GB.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold);
    $chkFilter1GB.Checked = $true;
    $pnlCol2.Controls.Add($chkFilter1GB);

    $chkScanInstallers = New-Object System.Windows.Forms.CheckBox;
    $chkScanInstallers.Text = "Setup-Dateien analysieren (.exe, .msi, .msp)";
    $chkScanInstallers.Location = New-Object System.Drawing.Point(8, 30);
    $chkScanInstallers.Size = New-Object System.Drawing.Size(340, 22);
    $chkScanInstallers.Checked = $true;
    $pnlCol2.Controls.Add($chkScanInstallers);

    $chkSkipSysFolders = New-Object System.Windows.Forms.CheckBox;
    $chkSkipSysFolders.Text = "Systemordner überspringen (Windows, Recovery...)";
    $chkSkipSysFolders.Location = New-Object System.Drawing.Point(8, 56);
    $chkSkipSysFolders.Size = New-Object System.Drawing.Size(380, 22);
    $chkSkipSysFolders.Checked = $true;
    $pnlCol2.Controls.Add($chkSkipSysFolders);

    # Col 3: Scan Button
    $pnlCol3 = New-Object System.Windows.Forms.Panel;
    $pnlCol3.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $tblTop.Controls.Add($pnlCol3, 3, 0);

    $btnScan = New-Object System.Windows.Forms.Button;
    $btnScan.Text = "Analyse starten";
    $btnScan.Location = New-Object System.Drawing.Point(20, 10);
    $btnScan.Size = New-Object System.Drawing.Size(180, 68);
    $btnScan.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215);
    $btnScan.ForeColor = [System.Drawing.Color]::White;
    $btnScan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat;
    $btnScan.Font = New-Object System.Drawing.Font("Segoe UI", 11.0, [System.Drawing.FontStyle]::Bold);
    $pnlCol3.Controls.Add($btnScan);

    # =========================================================================
    # STATUSLEISTE
    # =========================================================================
    $pnlStatus = New-Object System.Windows.Forms.Panel;
    $pnlStatus.Dock = [System.Windows.Forms.DockStyle]::Bottom;
    $pnlStatus.Height = 32;
    $pnlStatus.BackColor = [System.Drawing.Color]::FromArgb(235, 240, 248);
    $pnlStatus.Padding = New-Object System.Windows.Forms.Padding(12, 0, 12, 0);
    $form.Controls.Add($pnlStatus);

    $lblLiveDetail = New-Object System.Windows.Forms.Label;
    $lblLiveDetail.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $lblLiveDetail.Text = "Bereit. Analyse starten.";
    $lblLiveDetail.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft;
    $pnlStatus.Controls.Add($lblLiveDetail);

    # =========================================================================
    # TAB CONTROL (VIER REGISTERKARTEN)
    # =========================================================================
    $tabControl = New-Object System.Windows.Forms.TabControl;
    $tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9.5);
    $form.Controls.Add($tabControl);
    $tabControl.BringToFront();

    $tabOverview = New-Object System.Windows.Forms.TabPage;
    $tabOverview.Text = "  Übersicht & Wartungsbefehle  ";
    $tabOverview.BackColor = [System.Drawing.Color]::White;
    $tabControl.TabPages.Add($tabOverview);

    $tabDetails = New-Object System.Windows.Forms.TabPage;
    $tabDetails.Text = "  Detailanalyse (Hauptordner & Drilldown)  ";
    $tabDetails.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252);
    $tabControl.TabPages.Add($tabDetails);

    $tabUsers = New-Object System.Windows.Forms.TabPage;
    $tabUsers.Text = "  Benutzerprofile & Dateianalyse  ";
    $tabUsers.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252);
    $tabControl.TabPages.Add($tabUsers);

    $tabRecycle = New-Object System.Windows.Forms.TabPage;
    $tabRecycle.Text = "  Papierkorb ($Recycle.Bin & SIDs)  ";
    $tabRecycle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252);
    $tabControl.TabPages.Add($tabRecycle);

    # -------------------------------------------------------------------------
    # TAB 1: ÜBERSICHT, LAUFWERKE & WARTUNG
    # -------------------------------------------------------------------------
    $pnlOverviewScroll = New-Object System.Windows.Forms.Panel;
    $pnlOverviewScroll.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $pnlOverviewScroll.AutoScroll = $true;
    $pnlOverviewScroll.Padding = New-Object System.Windows.Forms.Padding(15, 12, 15, 15);
    $tabOverview.Controls.Add($pnlOverviewScroll);

    $lblDrivesHeader = New-Object System.Windows.Forms.Label;
    $lblDrivesHeader.Text = "Erkannte Laufwerke & Auslastung:";
    $lblDrivesHeader.Font = New-Object System.Drawing.Font("Segoe UI", 11.0, [System.Drawing.FontStyle]::Bold);
    $lblDrivesHeader.Location = New-Object System.Drawing.Point(15, 10);
    $lblDrivesHeader.AutoSize = $true;
    $pnlOverviewScroll.Controls.Add($lblDrivesHeader);

    $pnlDriveCards = New-Object System.Windows.Forms.FlowLayoutPanel;
    $pnlDriveCards.Location = New-Object System.Drawing.Point(15, 38);
    $pnlDriveCards.Size = New-Object System.Drawing.Size(1420, 175);
    $pnlDriveCards.AutoScroll = $true;
    $pnlOverviewScroll.Controls.Add($pnlDriveCards);

    $lblCleanupTitle = New-Object System.Windows.Forms.Label;
    $lblCleanupTitle.Text = "Bereinigungs-Befehle & Potenzielle Speicherplatz-Rückgewinnung:";
    $lblCleanupTitle.Font = New-Object System.Drawing.Font("Segoe UI", 11.0, [System.Drawing.FontStyle]::Bold);
    $lblCleanupTitle.Location = New-Object System.Drawing.Point(15, 225);
    $lblCleanupTitle.AutoSize = $true;
    $pnlOverviewScroll.Controls.Add($lblCleanupTitle);

    $gridCleanup = New-Object System.Windows.Forms.DataGridView;
    $gridCleanup.Location = New-Object System.Drawing.Point(15, 255);
    $gridCleanup.Size = New-Object System.Drawing.Size(1420, 360);
    $gridCleanup.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill;
    $pnlOverviewScroll.Controls.Add($gridCleanup);

    [void]$gridCleanup.Columns.Add("colArea", "Bereich");
    [void]$gridCleanup.Columns.Add("colCommand", "Ausgeführter Befehl");
    [void]$gridCleanup.Columns.Add("colSaved", "Geschätzte Ersparnis");
    [void]$gridCleanup.Columns.Add("colInfo", "Auswirkung");

    $gridCleanup.Columns["colArea"].FillWeight    = 45;
    $gridCleanup.Columns["colCommand"].FillWeight = 110;
    $gridCleanup.Columns["colSaved"].FillWeight   = 35;
    $gridCleanup.Columns["colInfo"].FillWeight    = 100;

    Apply-StandardGridTheme -grid $gridCleanup;

    $btnRunCleanupCmd = New-Object System.Windows.Forms.Button;
    $btnRunCleanupCmd.Text = "Markierten Befehl lokal ausführen";
    $btnRunCleanupCmd.Location = New-Object System.Drawing.Point(15, 628);
    $btnRunCleanupCmd.Size = New-Object System.Drawing.Size(250, 34);
    $btnRunCleanupCmd.BackColor = [System.Drawing.Color]::FromArgb(230, 240, 250);
    $btnRunCleanupCmd.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold);
    $btnRunCleanupCmd.FlatStyle = [System.Windows.Forms.FlatStyle]::System;
    $pnlOverviewScroll.Controls.Add($btnRunCleanupCmd);

    # -------------------------------------------------------------------------
    # TAB 2: DETAILANALYSE (HAUPTORDNER DRILLDOWN)
    # -------------------------------------------------------------------------
    $pnlFilter = New-Object System.Windows.Forms.Panel;
    $pnlFilter.Dock = [System.Windows.Forms.DockStyle]::Top;
    $pnlFilter.Height = 48;
    $pnlFilter.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250);
    $tabDetails.Controls.Add($pnlFilter);

    $lblSearch = New-Object System.Windows.Forms.Label;
    $lblSearch.Text = "Ordner filtern:";
    $lblSearch.Location = New-Object System.Drawing.Point(14, 14);
    $lblSearch.AutoSize = $true;
    $lblSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold);
    $pnlFilter.Controls.Add($lblSearch);

    $txtSearch = New-Object System.Windows.Forms.TextBox;
    $txtSearch.Location = New-Object System.Drawing.Point(115, 11);
    $txtSearch.Size = New-Object System.Drawing.Size(260, 25);
    $pnlFilter.Controls.Add($txtSearch);

    $btnExpandAll = New-Object System.Windows.Forms.Button;
    $btnExpandAll.Text = "Alle aufklappen";
    $btnExpandAll.Location = New-Object System.Drawing.Point(390, 9);
    $btnExpandAll.Size = New-Object System.Drawing.Size(120, 28);
    $btnExpandAll.FlatStyle = [System.Windows.Forms.FlatStyle]::System;
    $pnlFilter.Controls.Add($btnExpandAll);

    $btnCollapseAll = New-Object System.Windows.Forms.Button;
    $btnCollapseAll.Text = "Alle zuklappen";
    $btnCollapseAll.Location = New-Object System.Drawing.Point(520, 9);
    $btnCollapseAll.Size = New-Object System.Drawing.Size(120, 28);
    $btnCollapseAll.FlatStyle = [System.Windows.Forms.FlatStyle]::System;
    $pnlFilter.Controls.Add($btnCollapseAll);

    $grid = New-Object System.Windows.Forms.DataGridView;
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill;
    $tabDetails.Controls.Add($grid);
    $grid.BringToFront();

    [void]$grid.Columns.Add("colExp", "Tree");
    [void]$grid.Columns.Add("colDrive", "Laufwerk");
    [void]$grid.Columns.Add("colName", "Verzeichnis");
    [void]$grid.Columns.Add("colSizeGB", "Größe (GB)");
    [void]$grid.Columns.Add("colFiles", "Dateien");
    [void]$grid.Columns.Add("colInstallers", "Installer (.exe/.msi/.msp)");
    [void]$grid.Columns.Add("colInstSize", "Installer-Größe");
    [void]$grid.Columns.Add("colPath", "Pfad");

    $grid.Columns["colExp"].FillWeight        = 12;
    $grid.Columns["colDrive"].FillWeight      = 15;
    $grid.Columns["colName"].FillWeight       = 55;
    $grid.Columns["colSizeGB"].FillWeight     = 20;
    $grid.Columns["colFiles"].FillWeight      = 15;
    $grid.Columns["colInstallers"].FillWeight = 25;
    $grid.Columns["colInstSize"].FillWeight   = 22;
    $grid.Columns["colPath"].FillWeight       = 110;

    $grid.Columns["colExp"].DefaultCellStyle.Alignment = [System.Drawing.ContentAlignment]::MiddleCenter;
    $grid.Columns["colExp"].DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", 10.0, [System.Drawing.FontStyle]::Bold);

    Apply-StandardGridTheme -grid $grid;

    # -------------------------------------------------------------------------
    # TAB 3: BENUTZERPROFILE (.ost, > 1GB, .exe, .zip)
    # -------------------------------------------------------------------------
    $splitUsers = New-Object System.Windows.Forms.SplitContainer;
    $splitUsers.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $splitUsers.Orientation = [System.Windows.Forms.Orientation]::Horizontal;
    $splitUsers.SplitterDistance = 320;
    $splitUsers.SplitterWidth = 6;
    $tabUsers.Controls.Add($splitUsers);

    $pnlUserTop = New-Object System.Windows.Forms.Panel;
    $pnlUserTop.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $pnlUserTop.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 4);
    $splitUsers.Panel1.Controls.Add($pnlUserTop);

    $lblUserGridTitle = New-Object System.Windows.Forms.Label;
    $lblUserGridTitle.Text = "Erkannte Benutzerprofile (Klicken zur Anzeige der auffälligen Dateien):";
    $lblUserGridTitle.Dock = [System.Windows.Forms.DockStyle]::Top;
    $lblUserGridTitle.Height = 26;
    $lblUserGridTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10.0, [System.Drawing.FontStyle]::Bold);
    $pnlUserTop.Controls.Add($lblUserGridTitle);

    $gridUsers = New-Object System.Windows.Forms.DataGridView;
    $gridUsers.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $gridUsers.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill;
    $pnlUserTop.Controls.Add($gridUsers);
    $gridUsers.BringToFront();

    [void]$gridUsers.Columns.Add("colUserName", "Benutzerprofil");
    [void]$gridUsers.Columns.Add("colUserSizeGB", "Gesamtgröße (GB)");
    [void]$gridUsers.Columns.Add("colUserFiles", "Dateien gesamt");
    [void]$gridUsers.Columns.Add("colUserGT1GB", "Dateien > 1.0 GB");
    [void]$gridUsers.Columns.Add("colUserGT1GBSize", "Größe > 1.0 GB");
    [void]$gridUsers.Columns.Add("colUserOst", "Outlook (.ost / .pst)");
    [void]$gridUsers.Columns.Add("colUserOstSize", "Outlook-Größe");
    [void]$gridUsers.Columns.Add("colUserExe", "Installer (.exe / .msi)");
    [void]$gridUsers.Columns.Add("colUserZip", "Archive (.zip / .rar)");
    [void]$gridUsers.Columns.Add("colUserPath", "Profilpfad");

    $gridUsers.Columns["colUserName"].FillWeight     = 45;
    $gridUsers.Columns["colUserSizeGB"].FillWeight   = 35;
    $gridUsers.Columns["colUserFiles"].FillWeight    = 30;
    $gridUsers.Columns["colUserGT1GB"].FillWeight    = 35;
    $gridUsers.Columns["colUserGT1GBSize"].FillWeight= 35;
    $gridUsers.Columns["colUserOst"].FillWeight      = 40;
    $gridUsers.Columns["colUserOstSize"].FillWeight  = 35;
    $gridUsers.Columns["colUserExe"].FillWeight      = 40;
    $gridUsers.Columns["colUserZip"].FillWeight      = 35;
    $gridUsers.Columns["colUserPath"].FillWeight     = 90;

    Apply-StandardGridTheme -grid $gridUsers;

    $pnlUserBottom = New-Object System.Windows.Forms.Panel;
    $pnlUserBottom.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $pnlUserBottom.Padding = New-Object System.Windows.Forms.Padding(12, 4, 12, 8);
    $splitUsers.Panel2.Controls.Add($pnlUserBottom);

    $lblUserFileTitle = New-Object System.Windows.Forms.Label;
    $lblUserFileTitle.Text = "Auffällige Dateien im gewählten Profil (Doppelklick zum Öffnen im Explorer):";
    $lblUserFileTitle.Dock = [System.Windows.Forms.DockStyle]::Top;
    $lblUserFileTitle.Height = 26;
    $lblUserFileTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10.0, [System.Drawing.FontStyle]::Bold);
    $pnlUserBottom.Controls.Add($lblUserFileTitle);

    $gridUserFiles = New-Object System.Windows.Forms.DataGridView;
    $gridUserFiles.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $gridUserFiles.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill;
    $pnlUserBottom.Controls.Add($gridUserFiles);
    $gridUserFiles.BringToFront();

    [void]$gridUserFiles.Columns.Add("colFType", "Kategorie / Typ");
    [void]$gridUserFiles.Columns.Add("colFName", "Dateiname");
    [void]$gridUserFiles.Columns.Add("colFSizeMB", "Größe (MB)");
    [void]$gridUserFiles.Columns.Add("colFSizeGB", "Größe (GB)");
    [void]$gridUserFiles.Columns.Add("colFExt", "Endung");
    [void]$gridUserFiles.Columns.Add("colFPath", "Vollständiger Dateipfad");

    $gridUserFiles.Columns["colFType"].FillWeight   = 40;
    $gridUserFiles.Columns["colFName"].FillWeight   = 70;
    $gridUserFiles.Columns["colFSizeMB"].FillWeight = 25;
    $gridUserFiles.Columns["colFSizeGB"].FillWeight = 25;
    $gridUserFiles.Columns["colFExt"].FillWeight    = 20;
    $gridUserFiles.Columns["colFPath"].FillWeight   = 140;

    Apply-StandardGridTheme -grid $gridUserFiles;

    # -------------------------------------------------------------------------
    # TAB 4: PAPIERKORB & SID-ANALYSE ($Recycle.Bin)
    # -------------------------------------------------------------------------
    $pnlRecycleTop = New-Object System.Windows.Forms.Panel;
    $pnlRecycleTop.Dock = [System.Windows.Forms.DockStyle]::Top;
    $pnlRecycleTop.Height = 55;
    $pnlRecycleTop.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250);
    $pnlRecycleTop.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 10);
    $tabRecycle.Controls.Add($pnlRecycleTop);

    $lblRecycleDesc = New-Object System.Windows.Forms.Label;
    $lblRecycleDesc.Text = "Aufschlüsselung aller $Recycle.Bin-Ordner nach Windows-Benutzer-SIDs (TreeSize-Prinzip):";
    $lblRecycleDesc.Location = New-Object System.Drawing.Point(12, 16);
    $lblRecycleDesc.AutoSize = $true;
    $lblRecycleDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold);
    $pnlRecycleTop.Controls.Add($lblRecycleDesc);

    $btnPurgeSidRecycle = New-Object System.Windows.Forms.Button;
    $btnPurgeSidRecycle.Text = "Papierkorb für gewählte SID leeren";
    $btnPurgeSidRecycle.Location = New-Object System.Drawing.Point(680, 10);
    $btnPurgeSidRecycle.Size = New-Object System.Drawing.Size(250, 32);
    $btnPurgeSidRecycle.BackColor = [System.Drawing.Color]::FromArgb(254, 242, 242);
    $btnPurgeSidRecycle.ForeColor = [System.Drawing.Color]::DarkRed;
    $btnPurgeSidRecycle.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold);
    $btnPurgeSidRecycle.FlatStyle = [System.Windows.Forms.FlatStyle]::System;
    $pnlRecycleTop.Controls.Add($btnPurgeSidRecycle);

    $gridRecycle = New-Object System.Windows.Forms.DataGridView;
    $gridRecycle.Dock = [System.Windows.Forms.DockStyle]::Fill;
    $gridRecycle.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill;
    $tabRecycle.Controls.Add($gridRecycle);
    $gridRecycle.BringToFront();

    [void]$gridRecycle.Columns.Add("colRecDrive", "Laufwerk");
    [void]$gridRecycle.Columns.Add("colRecUser", "Zugeordneter Benutzer");
    [void]$gridRecycle.Columns.Add("colRecSid", "Sicherheitskennung (SID)");
    [void]$gridRecycle.Columns.Add("colRecSizeMB", "Größe (MB)");
    [void]$gridRecycle.Columns.Add("colRecSizeGB", "Größe (GB)");
    [void]$gridRecycle.Columns.Add("colRecFiles", "Gelöschte Objekte");
    [void]$gridRecycle.Columns.Add("colRecStatus", "Konto-Status");
    [void]$gridRecycle.Columns.Add("colRecPath", "Pfad im $Recycle.Bin");

    $gridRecycle.Columns["colRecDrive"].FillWeight  = 18;
    $gridRecycle.Columns["colRecUser"].FillWeight   = 55;
    $gridRecycle.Columns["colRecSid"].FillWeight    = 60;
    $gridRecycle.Columns["colRecSizeMB"].FillWeight = 25;
    $gridRecycle.Columns["colRecSizeGB"].FillWeight = 25;
    $gridRecycle.Columns["colRecFiles"].FillWeight  = 25;
    $gridRecycle.Columns["colRecStatus"].FillWeight = 40;
    $gridRecycle.Columns["colRecPath"].FillWeight   = 100;

    Apply-StandardGridTheme -grid $gridRecycle;

    # =========================================================================
    # LOGIK: WARTUNGSTABELLE TAB 1
    # =========================================================================
    $PopulateCleanupTable = {
        $gridCleanup.Rows.Clear();

        # Benutzer-Temp
        $tempMB = 0;
        try {
            $tmpFiles = [System.IO.Directory]::GetFiles($env:TEMP, "*", [System.IO.SearchOption]::AllDirectories);
            $sumB = [int64]0;
            for ($fIdx = 0; $fIdx -lt $tmpFiles.Count; $fIdx++) {
                try { $sumB += (New-Object System.IO.FileInfo($tmpFiles[$fIdx])).Length; } catch {}
            }
            $tempMB = [Math]::Round(($sumB / 1MB), 1);
        } catch {}

        [void]$gridCleanup.Rows.Add(
            "Benutzer Temp-Dateien",
            "Remove-Item `"$env:TEMP\*`" -Recurse -Force",
            "$tempMB MB",
            "Löscht temporäre App- und Cache-Rückstände des angemeldeten Benutzers."
        );

        # Adobe ARM
        $adobeArmMB = 0;
        $adobePaths = @(
            "$env:ProgramData\Adobe\ARM",
            "$env:LOCALAPPDATA\Adobe\ARM"
        );
        $adobeArmFound = $false;

        for ($p = 0; $p -lt $adobePaths.Count; $p++) {
            $armPath = $adobePaths[$p];
            if ([System.IO.Directory]::Exists($armPath)) {
                $adobeArmFound = $true;
                try {
                    $armFiles = [System.IO.Directory]::GetFiles($armPath, "*.*", [System.IO.SearchOption]::AllDirectories);
                    $sumArm = [int64]0;
                    for ($a = 0; $a -lt $armFiles.Count; $a++) {
                        try { $sumArm += (New-Object System.IO.FileInfo($armFiles[$a])).Length; } catch {}
                    }
                    $adobeArmMB += [Math]::Round(($sumArm / 1MB), 1);
                } catch {}
            }
        }

        $adobeCmd = "Stop-Service -Name AdobeARMservice -ErrorAction SilentlyContinue; Get-ChildItem -Path @('$env:ProgramData\Adobe\ARM', '$env:LOCALAPPDATA\Adobe\ARM') -Recurse -File -Include *.msi,*.msp,*.log | Remove-Item -Force";
        $adobeSaveText = if ($adobeArmFound) { "$adobeArmMB MB" } else { "Nicht installiert" };

        [void]$gridCleanup.Rows.Add(
            "Adobe ARM (Patch-Cache)",
            $adobeCmd,
            $adobeSaveText,
            "Stoppt den Updater-Dienst und entfernt kumulierte .msp/.msi Setup-Dateien und Logs."
        );

        # Windows Update (DISM)
        [void]$gridCleanup.Rows.Add(
            "Windows Update Bereinigung",
            "Dism.exe /online /Cleanup-Image /StartComponentCleanup /ResetBase",
            "ca. 2 - 8 GB",
            "Entfernt veraltete Komponentenversionen dauerhaft aus dem WinSxS-Verzeichnis."
        );

        # Datenträgerbereinigung
        [void]$gridCleanup.Rows.Add(
            "Datenträgerbereinigung",
            "cleanmgr.exe /sagerun:1",
            "ca. 1 - 5 GB",
            "Natives Bereinigungstool für Crashdumps, Miniaturansichten und Systemcaches."
        );

        # SoftwareDistribution
        [void]$gridCleanup.Rows.Add(
            "Windows Update Download Cache",
            "net stop wuauserv; Remove-Item C:\Windows\SoftwareDistribution\Download\* -Recurse",
            "ca. 1 - 4 GB",
            "Löscht heruntergeladene Installationspakete bereits angewendeter Updates."
        );

        # Schattenkopien
        [void]$gridCleanup.Rows.Add(
            "Schattenkopien (VSS)",
            "vssadmin delete shadows /for=C: /oldest",
            "ca. 3 - 15 GB",
            "Gibt Speicherplatz frei, der von alten System-Wiederherstellungspunkten belegt ist."
        );
    };

    # =========================================================================
    # LOGIK: LAUFWERKE LADEN
    # =========================================================================
    $LoadDrivesLogic = {
        $target = $txtHost.Text.Trim();
        if ([string]::IsNullOrWhiteSpace($target)) { $target = "localhost"; }

        $chkListDrives.Items.Clear();
        $pnlDriveCards.Controls.Clear();
        $script:DriveList.Clear();
        $lblLiveDetail.Text = "Lese Laufwerke von $target ein...";
        [System.Windows.Forms.Application]::DoEvents();

        try {
            $isLocal = ($target -in @("localhost", "127.0.0.1", $env:COMPUTERNAME));
            $drivesRaw = if ($isLocal) {
                Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType = 3" -ErrorAction Stop;
            } else {
                Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $target -Filter "DriveType = 3" -ErrorAction Stop;
            };

            $drivesArr = @($drivesRaw);
            for ($dIdx = 0; $dIdx -lt $drivesArr.Count; $dIdx++) {
                $d = $drivesArr[$dIdx];
                $totalGB = [Math]::Round(($d.Size / 1GB), 1);
                $freeGB  = [Math]::Round(($d.FreeSpace / 1GB), 1);
                $usedGB  = [Math]::Round(($totalGB - $freeGB), 1);
                $usedPct = if ($totalGB -gt 0) { [Math]::Round(($usedGB / $totalGB) * 100) } else { 0 };

                $chkListDrives.Items.Add("$($d.DeviceID) ($totalGB GB)", $true);

                $script:DriveList.Add([PSCustomObject]@{
                    DeviceID = $d.DeviceID;
                    TotalGB  = $totalGB;
                    FreeGB   = $freeGB;
                    UsedPct  = $usedPct;
                });

                $card = New-Object System.Windows.Forms.TableLayoutPanel;
                $card.Size = New-Object System.Drawing.Size(320, 150);
                $card.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle;
                $card.BackColor = [System.Drawing.Color]::White;
                $card.Margin = New-Object System.Windows.Forms.Padding(0, 0, 18, 12);
                $card.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 8);
                $card.ColumnCount = 1;
                $card.RowCount = 4;
                $card.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)));
                $card.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24)));
                $card.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24)));
                $card.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)));

                $lblD = New-Object System.Windows.Forms.Label;
                $lblD.Text = "Laufwerk $($d.DeviceID)  ($totalGB GB)";
                $lblD.Font = New-Object System.Drawing.Font("Segoe UI", 10.5, [System.Drawing.FontStyle]::Bold);
                $lblD.Dock = [System.Windows.Forms.DockStyle]::Fill;
                $card.Controls.Add($lblD, 0, 0);

                $lblBelegt = New-Object System.Windows.Forms.Label;
                $lblBelegt.Text = "Belegt: $usedGB GB ($usedPct%)";
                $lblBelegt.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold);
                $lblBelegt.ForeColor = if ($usedPct -ge 90) { [System.Drawing.Color]::DarkRed } else { [System.Drawing.Color]::FromArgb(30, 41, 59) };
                $lblBelegt.Dock = [System.Windows.Forms.DockStyle]::Fill;
                $card.Controls.Add($lblBelegt, 0, 1);

                $lblFrei = New-Object System.Windows.Forms.Label;
                $lblFrei.Text = "Frei:     $freeGB GB";
                $lblFrei.Font = New-Object System.Drawing.Font("Segoe UI", 9.0);
                $lblFrei.ForeColor = [System.Drawing.Color]::FromArgb(60, 70, 80);
                $lblFrei.Dock = [System.Windows.Forms.DockStyle]::Fill;
                $card.Controls.Add($lblFrei, 0, 2);

                $pb = New-Object System.Windows.Forms.ProgressBar;
                $pb.Dock = [System.Windows.Forms.DockStyle]::Fill;
                $pb.Value = [Math]::Min(100, [Math]::Max(0, [int]$usedPct));
                $card.Controls.Add($pb, 0, 3);

                $pnlDriveCards.Controls.Add($card);
            }
            $lblLiveDetail.Text = "Bereit. $($script:DriveList.Count) Laufwerk(e) geladen.";
        } catch {
            $lblLiveDetail.Text = "Fehler: $($_.Exception.Message)";
        }
    };

    $btnLoadDrives.Add_Click($LoadDrivesLogic);

    # Hilfsfunktion für Ordner-Metriken
    function Analyze-FolderMetrics {
        param([string]$Path, [bool]$CheckInstallers, $Fso)
        $res = @{
            SizeBytes     = [int64]0;
            Files         = 0;
            SubDirs       = 0;
            InstCount     = 0;
            InstSizeBytes = [int64]0;
        };

        if ($null -ne $Fso -and $Fso.FolderExists($Path)) {
            try {
                $fObj = $Fso.GetFolder($Path);
                $res.SizeBytes = [int64]$fObj.Size;
                $res.Files     = [int64]$fObj.Files.Count;
                $res.SubDirs   = [int64]$fObj.SubFolders.Count;
            } catch {}
        }

        if ($CheckInstallers -and [System.IO.Directory]::Exists($Path)) {
            try {
                $allFiles = [System.IO.Directory]::GetFiles($Path, "*.*", [System.IO.SearchOption]::AllDirectories);
                for ($i = 0; $i -lt $allFiles.Count; $i++) {
                    $ext = [System.IO.Path]::GetExtension($allFiles[$i]).ToLower();
                    if ($ext -in @(".exe", ".msi", ".msp")) {
                        $res.InstCount++;
                        try {
                            $fi = New-Object System.IO.FileInfo($allFiles[$i]);
                            $res.InstSizeBytes += $fi.Length;
                        } catch {}
                    }
                }
            } catch {}
        }
        return $res;
    }

    # =========================================================================
    # SCAN-FUNKTION: PAPIERKORB ($Recycle.Bin & SIDs - TAB 4)
    # =========================================================================
    $ScanRecycleBinLogic = {
        param($SelectedDrives, $TargetHost, $IsLocalHost, $Fso)

        $script:RecycleBinData.Clear();
        $gridRecycle.Rows.Clear();

        for ($d = 0; $d -lt $SelectedDrives.Count; $d++) {
            $drv = $SelectedDrives[$d];
            $drvLetter = $drv.Substring(0, 1);
            $recycleRoot = if ($IsLocalHost) { "$drvLetter`:\`$Recycle.Bin" } else { "\\$TargetHost\$drvLetter$\$Recycle.Bin" };

            if (-not [System.IO.Directory]::Exists($recycleRoot)) { continue; }

            $lblLiveDetail.Text = "Analysiere Papierkorb ($drv): $recycleRoot...";
            [System.Windows.Forms.Application]::DoEvents();

            $sidDirs = @();
            try { $sidDirs = [System.IO.Directory]::GetDirectories($recycleRoot); } catch {}

            for ($s = 0; $s -lt $sidDirs.Count; $s++) {
                $sDir = $sidDirs[$s];
                $sidString = [System.IO.Path]::GetFileName($sDir);

                # Nur gültige Sicherheitskennungen auswerten
                if ($sidString -notmatch '^S-1-5-') { continue; }

                # SID in Benutzername auflösen
                $accountName = "Unbekanntes Konto";
                $accountStatus = "Aktiv / Bekannt";

                try {
                    $sidObj = New-Object System.Security.Principal.SecurityIdentifier($sidString);
                    $ntAccount = $sidObj.Translate([System.Security.Principal.NTAccount]);
                    $accountName = $ntAccount.Value;
                } catch {
                    $accountName = "Gelöschtes Konto / Verwaiste SID";
                    $accountStatus = "Verwaist (Orphaned)";
                }

                # Größe & Dateien via FSO
                $binBytes = [int64]0;
                $binFiles = 0;
                if ($null -ne $Fso -and $Fso.FolderExists($sDir)) {
                    try {
                        $fObj = $Fso.GetFolder($sDir);
                        $binBytes = [int64]$fObj.Size;
                        $binFiles = [int64]$fObj.Files.Count;
                    } catch {}
                }

                $sizeMB = [Math]::Round(($binBytes / 1MB), 1);
                $sizeGB = [Math]::Round(($binBytes / 1GB), 2);

                $recItem = [PSCustomObject]@{
                    Drive     = $drv;
                    User      = $accountName;
                    SID       = $sidString;
                    SizeMB    = $sizeMB;
                    SizeGB    = $sizeGB;
                    Files     = $binFiles;
                    Status    = $accountStatus;
                    Path      = $sDir;
                };

                $script:RecycleBinData.Add($recItem);

                $rIdx = $gridRecycle.Rows.Add(
                    $recItem.Drive,
                    $recItem.User,
                    $recItem.SID,
                    $recItem.SizeMB,
                    $recItem.SizeGB,
                    $recItem.Files,
                    $recItem.Status,
                    $recItem.Path
                );

                if ($accountStatus -match "Verwaist") {
                    $gridRecycle.Rows[$rIdx].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 83, 9);
                }
                if ($recItem.SizeGB -ge 1.0) {
                    $gridRecycle.Rows[$rIdx].Cells["colRecSizeGB"].Style.ForeColor = [System.Drawing.Color]::DarkRed;
                    $gridRecycle.Rows[$rIdx].Cells["colRecSizeGB"].Style.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold);
                }
            }
        }
    };

    # Papierkorb für die gewählte SID leeren
    $btnPurgeSidRecycle.Add_Click({
        if ($gridRecycle.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie zuerst eine SID-Zeile in der Tabelle aus.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information);
            return;
        }

        $selRow = $gridRecycle.SelectedRows[0];
        $targetPath = [string]$selRow.Cells["colRecPath"].Value;
        $targetUser = [string]$selRow.Cells["colRecUser"].Value;
        $targetSid  = [string]$selRow.Cells["colRecSid"].Value;

        $msg = "Möchten Sie den Papierkorb für folgenden Benutzer wirklich unwiderruflich leeren?`n`nBenutzer: $targetUser`nSID: $targetSid`nPfad: $targetPath";
        $confirm = [System.Windows.Forms.MessageBox]::Show($msg, "Papierkorb-Bereinigung bestätigen", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning);

        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                $purgeCmd = "Remove-Item -Path `"$targetPath\*`" -Recurse -Force -ErrorAction Stop";
                Start-Process powershell.exe -ArgumentList "-NoProfile -Command `"$purgeCmd`"" -Verb RunAs -Wait;
                [System.Windows.Forms.MessageBox]::Show("Bereinigung ausgeführt!", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information);
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Leeren: $($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error);
            }
        }
    });

    # =========================================================================
    # SCAN-FUNKTION: BENUTZERPROFILE (TAB 3)
    # =========================================================================
    $ScanUserProfilesLogic = {
        param($TargetHost, $IsLocalHost, $Fso)

        $script:UserProfiles.Clear();
        $script:UserFilesMap.Clear();
        $gridUsers.Rows.Clear();
        $gridUserFiles.Rows.Clear();

        $usersBaseDir = if ($IsLocalHost) { "C:\Users" } else { "\\$TargetHost\C$\Users" };
        if (-not [System.IO.Directory]::Exists($usersBaseDir)) { return; }

        $excludedProfiles = @("All Users", "Default", "Default User", "Public", "desktop.ini");
        $userDirs = [System.IO.Directory]::GetDirectories($usersBaseDir);

        for ($u = 0; $u -lt $userDirs.Count; $u++) {
            $uDir = $userDirs[$u];
            $uName = [System.IO.Path]::GetFileName($uDir);
            if ($uName -in $excludedProfiles) { continue; }

            $lblLiveDetail.Text = "Analysiere Benutzerprofil: $uName...";
            [System.Windows.Forms.Application]::DoEvents();

            $uSize = [int64]0;
            $uFiles = 0;
            if ($null -ne $Fso -and $Fso.FolderExists($uDir)) {
                try {
                    $fObj = $Fso.GetFolder($uDir);
                    $uSize = [int64]$fObj.Size;
                    $uFiles = [int64]$fObj.Files.Count;
                } catch {}
            }

            $countGT1GB = 0;
            $sizeGT1GB = [int64]0;
            $countOst = 0;
            $sizeOst = [int64]0;
            $countExe = 0;
            $sizeExe = [int64]0;
            $countZip = 0;
            $sizeZip = [int64]0;

            $fileDetailsList = [System.Collections.Generic.List[PSCustomObject]]::new();

            try {
                $rawFiles = [System.IO.Directory]::GetFiles($uDir, "*.*", [System.IO.SearchOption]::AllDirectories);
                for ($f = 0; $f -lt $rawFiles.Count; $f++) {
                    $filePath = $rawFiles[$f];
                    $fi = $null;
                    try { $fi = New-Object System.IO.FileInfo($filePath); } catch { continue; }
                    if ($null -eq $fi) { continue; }

                    $ext = $fi.Extension.ToLower();
                    $fLen = $fi.Length;
                    $isInteresting = $false;
                    $category = "";

                    if ($fLen -ge 1GB) {
                        $countGT1GB++;
                        $sizeGT1GB += $fLen;
                        $isInteresting = $true;
                        $category = "Datei > 1.0 GB";
                    }

                    if ($ext -in @(".ost", ".pst")) {
                        $countOst++;
                        $sizeOst += $fLen;
                        $isInteresting = $true;
                        if ([string]::IsNullOrWhiteSpace($category)) { $category = "Outlook Datendatei"; } else { $category += " / Outlook"; }
                    }

                    if ($ext -in @(".exe", ".msi", ".msp")) {
                        $countExe++;
                        $sizeExe += $fLen;
                        if ($fLen -ge 20MB) {
                            $isInteresting = $true;
                            if ([string]::IsNullOrWhiteSpace($category)) { $category = "Setup / Installer"; } else { $category += " / Setup"; }
                        }
                    }

                    if ($ext -in @(".zip", ".rar", ".7z", ".tar", ".gz")) {
                        $countZip++;
                        $sizeZip += $fLen;
                        if ($fLen -ge 50MB) {
                            $isInteresting = $true;
                            if ([string]::IsNullOrWhiteSpace($category)) { $category = "Archiv"; } else { $category += " / Archiv"; }
                        }
                    }

                    if ($isInteresting) {
                        $fileDetailsList.Add([PSCustomObject]@{
                            Category = $category;
                            FileName = $fi.Name;
                            SizeMB   = [Math]::Round(($fLen / 1MB), 1);
                            SizeGB   = [Math]::Round(($fLen / 1GB), 2);
                            Ext      = $ext;
                            FullPath = $fi.FullName;
                        });
                    }
                }
            } catch {}

            $uObj = [PSCustomObject]@{
                UserName       = $uName;
                TotalSizeGB    = [Math]::Round(($uSize / 1GB), 2);
                TotalFiles     = $uFiles;
                CountGT1GB     = $countGT1GB;
                SizeGT1GBStr   = "$([Math]::Round(($sizeGT1GB / 1GB), 2)) GB";
                CountOst       = $countOst;
                SizeOstStr     = "$([Math]::Round(($sizeOst / 1GB), 2)) GB";
                CountExe       = $countExe;
                CountZip       = $countZip;
                ProfilePath    = $uDir;
            };

            $script:UserProfiles.Add($uObj);
            $script:UserFilesMap[$uName] = $fileDetailsList;

            $rIdx = $gridUsers.Rows.Add(
                $uObj.UserName,
                $uObj.TotalSizeGB,
                $uObj.TotalFiles,
                $uObj.CountGT1GB,
                $uObj.SizeGT1GBStr,
                $uObj.CountOst,
                $uObj.SizeOstStr,
                $uObj.CountExe,
                $uObj.CountZip,
                $uObj.ProfilePath
            );

            if ($uObj.CountGT1GB -gt 0 -or $uObj.TotalSizeGB -gt 25.0) {
                $gridUsers.Rows[$rIdx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(254, 243, 242);
                $gridUsers.Rows[$rIdx].Cells["colUserGT1GB"].Style.ForeColor = [System.Drawing.Color]::DarkRed;
                $gridUsers.Rows[$rIdx].Cells["colUserGT1GB"].Style.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold);
            }
        }
    };

    $gridUsers.Add_SelectionChanged({
        if ($gridUsers.SelectedRows.Count -eq 0) { return; }
        $selUser = [string]$gridUsers.SelectedRows[0].Cells["colUserName"].Value;
        $gridUserFiles.Rows.Clear();

        if ($selUser -and $script:UserFilesMap.ContainsKey($selUser)) {
            $files = $script:UserFilesMap[$selUser];
            $lblUserFileTitle.Text = "Auffällige Dateien im Profil [$selUser] ($($files.Count) Treffer gefunden):";

            for ($f = 0; $f -lt $files.Count; $f++) {
                $file = $files[$f];
                $fRow = $gridUserFiles.Rows.Add(
                    $file.Category,
                    $file.FileName,
                    $file.SizeMB,
                    $file.SizeGB,
                    $file.Ext,
                    $file.FullPath
                );

                if ($file.SizeGB -ge 1.0) {
                    $gridUserFiles.Rows[$fRow].Cells["colFSizeGB"].Style.ForeColor = [System.Drawing.Color]::DarkRed;
                    $gridUserFiles.Rows[$fRow].Cells["colFSizeGB"].Style.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold);
                }
            }
        }
    });

    # =========================================================================
    # REFRESH GRID TAB 2 (DRILLDOWN)
    # =========================================================================
    $RenderGridRows = {$grid.Rows.Clear();
        $term =$txtSearch.Text.Trim().ToLower();
        $only1GB =$chkFilter1GB.Checked;

        for ($rIdx = 0; $rIdx -lt$script:MasterFolders.Count; $rIdx++) {$root = $script:MasterFolders[$rIdx];

            if ($only1GB -and$root.SizeGB -lt 1.0) { continue; }

            $matchTerm = [string]::IsNullOrWhiteSpace($term) -or $root.Name.ToLower().Contains($term) -or $root.Path.ToLower().Contains($term);
            if (-not $matchTerm) { continue; }

            $isExpanded = $script:ExpandedRoots.Contains($root.Path);
            $hasSubs =$script:SubFoldersMap.ContainsKey($root.Path) -and ($script:SubFoldersMap[$root.Path].Count -gt 0);$treeSymbol = if ($hasSubs) { if ($isExpanded) { "[-]" } else { "[+]" } } else { " • " };

            $instSizeStr = if ($root.InstSizeBytes -gt 0) { "$([Math]::Round(($root.InstSizeBytes / 1MB), 1)) MB" } else { "-" };
            $instCountStr = if ($root.InstCount -gt 0) { "$($root.InstCount) Dateien" } else { "-" };

            $rowIdx =$grid.Rows.Add(
                $treeSymbol,$root.Drive,
                $root.Name,
                $root.SizeGB,
                $root.Files,
                $instCountStr,
                $instSizeStr,$root.Path
            );

            $grid.Rows[$rowIdx].DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Bold);
            $grid.Rows[$rowIdx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 246, 255);

            if ($isExpanded -and $hasSubs) {$subList = $script:SubFoldersMap[$root.Path];
                for ($sIdx = 0; $sIdx -lt$subList.Count; $sIdx++) {$sub = $subList[$sIdx];
                    if ($only1GB -and$sub.SizeGB -lt 1.0) { continue; }

                    $subInstSizeStr = if ($sub.InstSizeBytes -gt 0) { "$([Math]::Round(($sub.InstSizeBytes / 1MB), 1)) MB" } else { "-" };
                    $subInstCountStr = if ($sub.InstCount -gt 0) { "$($sub.InstCount) Dateien" } else { "-" };

                    $subRowIdx =$grid.Rows.Add(
                        "    ",
                        "",
                        ("  └── " + $sub.Name),$sub.SizeGB,
                        $sub.Files,
                        $subInstCountStr,
                        $subInstSizeStr,$sub.Path
                    );

                    $grid.Rows[$subRowIdx].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60);
                }
            }
        }
    };

    # =========================================================================
    # ANALYSE STARTEN
    # =========================================================================
    $btnScan.Add_Click({
        if ($script:IsScanning) { return; }

        $selectedDrives = [System.Collections.Generic.List[string]]::new();
        for ($cIdx = 0; $cIdx -lt$chkListDrives.CheckedItems.Count; $cIdx++) {$ci = $chkListDrives.CheckedItems[$cIdx];
            if ($ci -match '^([A-Za-z]:)') { $selectedDrives.Add($Matches[1]); }
        }

        if ($selectedDrives.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte mindestens ein Laufwerk ankreuzen.", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning);
            return;
        }

        $script:IsScanning =$true;
        $btnScan.Enabled =$false;
        $target =$txtHost.Text.Trim();
        if ([string]::IsNullOrWhiteSpace($target)) {$target = "localhost"; }
        $isLocal = ($target -in @("localhost", "127.0.0.1", $env:COMPUTERNAME));

        $script:MasterFolders.Clear();$script:SubFoldersMap.Clear();
        $script:ExpandedRoots.Clear();$grid.Rows.Clear();

        $fso =$null;
        try {
            $fso = New-Object -ComObject Scripting.FileSystemObject;
        } catch {}

        try {
            # 1. Haupt- und Unterordner analysieren (Tab 2)
            for ($drIdx = 0; $drIdx -lt $selectedDrives.Count; $drIdx++) {
                $drv =$selectedDrives[$drIdx];$drvLetter = $drv.Substring(0, 1);$rootPath = if ($isLocal) { "$drvLetter`:\" } else { "\\$target\$drvLetter$" };
                if (-not [System.IO.Directory]::Exists($rootPath)) { continue; }

                $rootInfo = [System.IO.DirectoryInfo]::new($rootPath);
                $level1Dirs = $rootInfo.GetDirectories();

                for ($l1Idx = 0; $l1Idx -lt $level1Dirs.Count; $l1Idx++) {
                    $l1 = $level1Dirs[$l1Idx];
                    if ($chkSkipSysFolders.Checked -and ($l1.Name -in @('$Recycle.Bin', 'System Volume Information', 'Windows'))) {
                        continue;
                    }

                    $lblLiveDetail.Text = "Scanne Hauptordner: $drv -> $($l1.Name)...";
                    [System.Windows.Forms.Application]::DoEvents();

                    $metrics1 = Analyze-FolderMetrics -Path $l1.FullName -CheckInstallers $chkScanInstallers.Checked -Fso $fso;
                    $sizeGB1 = [Math]::Round(($metrics1.SizeBytes / 1GB), 2);

                    $rootObj = [PSCustomObject]@{
                        Drive         = $drv;
                        Name          = $l1.Name;
                        Path          = $l1.FullName;
                        SizeGB        = $sizeGB1;
                        Files         = $metrics1.Files;
                        InstCount     = $metrics1.InstCount;
                        InstSizeBytes = $metrics1.InstSizeBytes;
                    };
                    $script:MasterFolders.Add($rootObj);

                    $script:SubFoldersMap[$l1.FullName] = [System.Collections.Generic.List[PSCustomObject]]::new();

                    $level2Dirs = @();
                    try { $level2Dirs = $l1.GetDirectories(); } catch {}

                    for ($l2Idx = 0; $l2Idx -lt $level2Dirs.Count; $l2Idx++) {
                        $l2 = $level2Dirs[$l2Idx];
                        $metrics2 = Analyze-FolderMetrics -Path $l2.FullName -CheckInstallers $chkScanInstallers.Checked -Fso $fso;
                        $sizeGB2 = [Math]::Round(($metrics2.SizeBytes / 1GB), 2);

                        $subObj = [PSCustomObject]@{
                            Name          = $l2.Name;
                            Path          = $l2.FullName;
                            SizeGB        = $sizeGB2;
                            Files         = $metrics2.Files;
                            InstCount     = $metrics2.InstCount;
                            InstSizeBytes = $metrics2.InstSizeBytes;
                        };
                        $script:SubFoldersMap[$l1.FullName].Add($subObj);
                    }
                }
            }

            & $RenderGridRows;

            # 2. Benutzerprofile & Dateianalyse starten (Tab 3)
            & $ScanUserProfilesLogic -TargetHost $target -IsLocalHost $isLocal -Fso $fso;

            # 3. Papierkorb & SID-Analyse starten (Tab 4)
            & $ScanRecycleBinLogic -SelectedDrives $selectedDrives -TargetHost $target -IsLocalHost $isLocal -Fso $fso;

            $lblLiveDetail.Text = "Analyse fertig: Hauptordner, Profile und Papierkorb-SIDs ausgewertet.";
            $tabControl.SelectedTab = $tabDetails;
        } finally {
            if ($null -ne $fso) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($fso); }
            $btnScan.Enabled = $true;
            $script:IsScanning = $false;
        }
    });

    # Klick-Event Tab 2 (Drilldown)
    $grid.Add_CellClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return; }

        $clickedPath = [string]$grid.Rows[$e.RowIndex].Cells["colPath"].Value;
        $treeSymbol = [string]$grid.Rows[$e.RowIndex].Cells["colExp"].Value;

        if ($treeSymbol -eq "[+]" -or $treeSymbol -eq "[-]") {
            if ($script:ExpandedRoots.Contains($clickedPath)) {
                [void]$script:ExpandedRoots.Remove($clickedPath);
            } else {
                [void]$script:ExpandedRoots.Add($clickedPath);
            }
            & $RenderGridRows;
        }
    });

    # Doppelklick auf Pfade
    $grid.Add_CellDoubleClick({
        param($s, $e)
        if ($e.RowIndex -ge 0) {
            $p = [string]$grid.Rows[$e.RowIndex].Cells["colPath"].Value;
            if ($p -and (Test-Path $p)) { Start-Process "explorer.exe" -ArgumentList "`"$p`""; }
        }
    });

    $gridRecycle.Add_CellDoubleClick({
        param($s, $e)
        if ($e.RowIndex -ge 0) {
            $p = [string]$gridRecycle.Rows[$e.RowIndex].Cells["colRecPath"].Value;
            if ($p -and (Test-Path $p)) { Start-Process "explorer.exe" -ArgumentList "`"$p`""; }
        }
    });

    $btnExpandAll.Add_Click({
        for ($i = 0; $i -lt $script:MasterFolders.Count; $i++) {
            [void]$script:ExpandedRoots.Add($script:MasterFolders[$i].Path);
        }
        & $RenderGridRows;
    });

    $btnCollapseAll.Add_Click({
        $script:ExpandedRoots.Clear();
        & $RenderGridRows;
    });

    $txtSearch.Add_KeyUp({ & $RenderGridRows; });
    $chkFilter1GB.Add_CheckedChanged({ & $RenderGridRows; });

    # Wartungsbefehl ausführen (Tab 1)
    $btnRunCleanupCmd.Add_Click({
        if ($gridCleanup.SelectedRows.Count -eq 0) { return; }
        $cmd = [string]$gridCleanup.SelectedRows[0].Cells["colCommand"].Value;
        $area = [string]$gridCleanup.SelectedRows[0].Cells["colArea"].Value;

        $ask = [System.Windows.Forms.MessageBox]::Show("Möchten Sie folgende Aktion jetzt mit Administratorrechten ausführen?`n`n$area`nBefehl: $cmd", "Bereinigung ausführen", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question);
        if ($ask -eq [System.Windows.Forms.DialogResult]::Yes) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -Command `"$cmd`"" -Verb RunAs;
        }
    });

    # Start-Initialisierung
    $form.Add_Shown({
        & $LoadDrivesLogic;
        & $PopulateCleanupTable;
    });

    try {
        [void]$form.ShowDialog();
    } finally {
        if ($null -ne $grid) {$grid.Dispose(); }
        if ($null -ne $gridCleanup) {$gridCleanup.Dispose(); }
        if ($null -ne $gridUsers) {$gridUsers.Dispose(); }
        if ($null -ne $gridUserFiles) {$gridUserFiles.Dispose(); }
        if ($null -ne $gridRecycle) {$gridRecycle.Dispose(); }
        if ($null -ne $form) {$form.Dispose(); }
        [System.GC]::Collect();
    }
}

Show-StorageAnalyzerGUI;
