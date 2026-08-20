<#
================================================================================
 GUI: HAUPTFENSTER, 3-SPALTEN HEADER & DASHBOARD
================================================================================
#>
function Start-AdminSuiteMainWindow {
    $mainForm = New-Object System.Windows.Forms.Form
    $mainForm.Text = Get-Text "Title"
    $mainForm.Size = New-Object System.Drawing.Size(980, 960)
    $mainForm.StartPosition = "CenterScreen"
    $mainForm.FormBorderStyle = "FixedDialog"
    $mainForm.MaximizeBox = $false
    $mainForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $mainForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # --- HEADER PANEL (3 Spalten) ---
    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlHeader.Height = 220
    $pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(235, 242, 250)
    $mainForm.Controls.Add($pnlHeader)

    # Sprachumschalter (DE / EN)
    $btnLangEN = New-Object System.Windows.Forms.Button
    $btnLangEN.Location = New-Object System.Drawing.Point(860, 10)
    $btnLangEN.Size = New-Object System.Drawing.Size(45, 26)
    $btnLangEN.Text = "EN"
    $btnLangEN.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8, [System.Drawing.FontStyle]::Bold)
    $pnlHeader.Controls.Add($btnLangEN)

    $btnLangDE = New-Object System.Windows.Forms.Button
    $btnLangDE.Location = New-Object System.Drawing.Point(910, 10)
    $btnLangDE.Size = New-Object System.Drawing.Size(45, 26)
    $btnLangDE.Text = "DE"
    $btnLangDE.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8, [System.Drawing.FontStyle]::Bold)
    $pnlHeader.Controls.Add($btnLangDE)

    # Hauptüberschrift im Header
    $lblHeaderMain = New-Object System.Windows.Forms.Label
    $lblHeaderMain.Location = New-Object System.Drawing.Point(18, 10)
    $lblHeaderMain.Size = New-Object System.Drawing.Size(830, 24)
    $lblHeaderMain.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 11.5, [System.Drawing.FontStyle]::Bold)
    $lblHeaderMain.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $lblHeaderMain.Text = "💻 $localComputerName | $localUserName"
    $pnlHeader.Controls.Add($lblHeaderMain)

    # Trennlinie
    $lblHeaderLine = New-Object System.Windows.Forms.Label
    $lblHeaderLine.Location = New-Object System.Drawing.Point(18, 38)
    $lblHeaderLine.Size = New-Object System.Drawing.Size(935, 1)
    $lblHeaderLine.BackColor = [System.Drawing.Color]::FromArgb(203, 213, 225)
    $pnlHeader.Controls.Add($lblHeaderLine)

    # -------------------------------------------------------------
    # SPALTE 1: Betriebssystem & Domäne (Links, X = 18, Breite = 300)
    # -------------------------------------------------------------
    $lblCol1Title = New-Object System.Windows.Forms.Label
    $lblCol1Title.Location = New-Object System.Drawing.Point(18, 46)
    $lblCol1Title.Size = New-Object System.Drawing.Size(300, 18)
    $lblCol1Title.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $lblCol1Title.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $pnlHeader.Controls.Add($lblCol1Title)

    $lblCol1Content = New-Object System.Windows.Forms.Label
    $lblCol1Content.Location = New-Object System.Drawing.Point(18, 66)
    $lblCol1Content.Size = New-Object System.Drawing.Size(300, 142)
    $lblCol1Content.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $lblCol1Content.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $pnlHeader.Controls.Add($lblCol1Content)

    # -------------------------------------------------------------
    # SPALTE 2: System & Hardware (Mitte, X = 330, Breite = 295)
    # -------------------------------------------------------------
    $lblCol2Title = New-Object System.Windows.Forms.Label
    $lblCol2Title.Location = New-Object System.Drawing.Point(330, 46)
    $lblCol2Title.Size = New-Object System.Drawing.Size(295, 18)
    $lblCol2Title.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $lblCol2Title.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $pnlHeader.Controls.Add($lblCol2Title)

    $lblCol2Content = New-Object System.Windows.Forms.Label
    $lblCol2Content.Location = New-Object System.Drawing.Point(330, 66)
    $lblCol2Content.Size = New-Object System.Drawing.Size(295, 142)
    $lblCol2Content.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $lblCol2Content.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $pnlHeader.Controls.Add($lblCol2Content)

    # -------------------------------------------------------------
    # SPALTE 3: Entra ID / Cloud Status (Rechts, X = 640, Breite = 315)
    # -------------------------------------------------------------
    $lblCol3Title = New-Object System.Windows.Forms.Label
    $lblCol3Title.Location = New-Object System.Drawing.Point(640, 46)
    $lblCol3Title.Size = New-Object System.Drawing.Size(315, 18)
    $lblCol3Title.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $lblCol3Title.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $pnlHeader.Controls.Add($lblCol3Title)

    $lblCol3Content = New-Object System.Windows.Forms.Label
    $lblCol3Content.Location = New-Object System.Drawing.Point(640, 66)
    $lblCol3Content.Size = New-Object System.Drawing.Size(315, 142)
    $lblCol3Content.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $lblCol3Content.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $pnlHeader.Controls.Add($lblCol3Content)

    # --- CLIENT TOOLS GROUPBOX ---
    $grpClient = New-Object System.Windows.Forms.GroupBox
    $grpClient.Location = New-Object System.Drawing.Point(18, 230)
    $grpClient.Size = New-Object System.Drawing.Size(935, 250)
    $grpClient.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $mainForm.Controls.Add($grpClient)

    $btnTool3 = New-Object System.Windows.Forms.Button; $btnTool3.Location = "20, 22"; $btnTool3.Size = "895, 46"
    $btnTool3.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool3.Add_Click({ Open-ToolEntraStatus }); $grpClient.Controls.Add($btnTool3)

    $btnTool4 = New-Object System.Windows.Forms.Button; $btnTool4.Location = "20, 74"; $btnTool4.Size = "895, 46"
    $btnTool4.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool4.Add_Click({ Open-ToolGroupsAndGPO }); $grpClient.Controls.Add($btnTool4)

    $btnTool5 = New-Object System.Windows.Forms.Button; $btnTool5.Location = "20, 126"; $btnTool5.Size = "895, 46"
    $btnTool5.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool5.Add_Click({ Open-ToolWin11Check }); $grpClient.Controls.Add($btnTool5)

    $btnTool8 = New-Object System.Windows.Forms.Button; $btnTool8.Location = "20, 178"; $btnTool8.Size = "895, 46"
    $btnTool8.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool8.Add_Click({ Show-ClientSoftwareAnalysis }); $grpClient.Controls.Add($btnTool8)

    # --- AD TOOLS GROUPBOX ---
    $grpAD = New-Object System.Windows.Forms.GroupBox
    $grpAD.Location = New-Object System.Drawing.Point(18, 490)
    $grpAD.Size = New-Object System.Drawing.Size(935, 410)
    $grpAD.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $mainForm.Controls.Add($grpAD)

    $btnTool1 = New-Object System.Windows.Forms.Button; $btnTool1.Location = "20, 22"; $btnTool1.Size = "895, 46"
    $btnTool1.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool1.Add_Click({ Open-ToolLastLogon }); $grpAD.Controls.Add($btnTool1)

    $btnTool2 = New-Object System.Windows.Forms.Button; $btnTool2.Location = "20, 74"; $btnTool2.Size = "895, 46"
    $btnTool2.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool2.Add_Click({ Open-ToolADAudit }); $grpAD.Controls.Add($btnTool2)

    $btnTool6 = New-Object System.Windows.Forms.Button; $btnTool6.Location = "20, 126"; $btnTool6.Size = "895, 46"
    $btnTool6.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool6.Add_Click({ Open-ToolDomainOverview }); $grpAD.Controls.Add($btnTool6)

    $btnTool7 = New-Object System.Windows.Forms.Button; $btnTool7.Location = "20, 178"; $btnTool7.Size = "895, 46"
    $btnTool7.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool7.Add_Click({ Open-ToolOSSupportAudit }); $grpAD.Controls.Add($btnTool7)

    $btnTool9 = New-Object System.Windows.Forms.Button; $btnTool9.Location = "20, 230"; $btnTool9.Size = "895, 46"
    $btnTool9.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool9.Add_Click({ Show-Tool9-ACLCompare }); $grpAD.Controls.Add($btnTool9)

    $btnTool10 = New-Object System.Windows.Forms.Button; $btnTool10.Location = "20, 282"; $btnTool10.Size = "895, 46"
    $btnTool10.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool10.Add_Click({ Show-Tool10-OUGroupFinder }); $grpAD.Controls.Add($btnTool10)

    $btnTool11 = New-Object System.Windows.Forms.Button; $btnTool11.Location = "20, 334"; $btnTool11.Size = "895, 46"
    $btnTool11.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool11.Add_Click({ Show-Tool11-PasswordPolicyAudit }); $grpAD.Controls.Add($btnTool11)

    # --- UI REFRESH FUNKTION ---
    function Update-UI {
        $mainForm.Text  = Get-Text "Title"
        $grpClient.Text = Get-Text "CategoryClient"
        $grpAD.Text     = Get-Text "CategoryAD"

        $lblCol1Title.Text = Get-Text "LblHdrOS"
        $lblCol2Title.Text = Get-Text "LblHdrSystem"
        $lblCol3Title.Text = Get-Text "LblHdrEntra"

        # Spalte 1: OS & Domäne
        $lblCol1Content.Text = @"
$("{0,-12}: {1}" -f (Get-Text "LblOS"), $osCaption)
$("{0,-12}: {1}" -f (Get-Text "LblBuild"), $osBuildNumber)
$("{0,-12}: {1}" -f (Get-Text "LblVersion"), $osVersionDisplay)
$("{0,-12}: {1}" -f (Get-Text "LblDomain"), $localDomainName)
$("{0,-12}: {1}" -f (Get-Text "LblLogonServer"), $localLogonServer)
"@

        # Spalte 2: System & Hardware
        $lblCol2Content.Text = @"
$("{0,-13}: {1}" -f (Get-Text "LblCompName"), $localComputerName)
$("{0,-13}: {1}" -f (Get-Text "LblManuf"), $localManufacturer)
$("{0,-13}: {1}" -f (Get-Text "LblModel"), $localModel)
$("{0,-13}: {1}" -f (Get-Text "LblSerial"), $localSerial)
$("{0,-13}: {1}" -f (Get-Text "LblSysType"), $localSystemType)
"@

        # Spalte 3: Entra ID / Cloud Status
        $lblCol3Content.Text = @"
$("{0,-12}: {1}" -f (Get-Text "LblJoinStatus"), $localJoinStatus)
$("{0,-12}: {1}" -f (Get-Text "LblPrtStatus"), $localAzureAdPrt)
$("{0,-12}: {1}" -f (Get-Text "LblTenantName"), $localTenantName)
$("{0,-12}: {1}" -f (Get-Text "LblTenantId"), $(if ($localTenantId.Length -gt 16) { $localTenantId.Substring(0,13) + "..." } else { $localTenantId }))
$("{0,-12}: {1}" -f (Get-Text "LblDeviceId"), $(if ($localDeviceId.Length -gt 16) { $localDeviceId.Substring(0,13) + "..." } else { $localDeviceId }))
"@

        $btnTool1.Text  = Get-Text "BtnTool1"
        $btnTool2.Text  = Get-Text "BtnTool2"
        $btnTool3.Text  = Get-Text "BtnTool3"
        $btnTool4.Text  = Get-Text "BtnTool4"
        $btnTool5.Text  = Get-Text "BtnTool5"
        $btnTool6.Text  = Get-Text "BtnTool6"
        $btnTool7.Text  = Get-Text "BtnTool7"
        $btnTool8.Text  = Get-Text "BtnTool8"
        $btnTool9.Text  = Get-Text "BtnTool9"
        $btnTool10.Text = Get-Text "BtnTool10"
        $btnTool11.Text = Get-Text "BtnTool11"

        $btnLangDE.BackColor = if ($script:CurrentLang -eq "DE") { [System.Drawing.Color]::FromArgb(37, 99, 235) } else { [System.Drawing.Color]::LightSteelBlue }
        $btnLangDE.ForeColor = if ($script:CurrentLang -eq "DE") { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black }
        $btnLangEN.BackColor = if ($script:CurrentLang -eq "EN") { [System.Drawing.Color]::FromArgb(37, 99, 235) } else { [System.Drawing.Color]::LightSteelBlue }
        $btnLangEN.ForeColor = if ($script:CurrentLang -eq "EN") { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black }
    }

    $btnLangEN.Add_Click({ $script:CurrentLang = "EN"; Update-UI })
    $btnLangDE.Add_Click({ $script:CurrentLang = "DE"; Update-UI })

    # Starten
    Update-UI
    [void]$mainForm.ShowDialog()
}
