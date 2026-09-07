<#
================================================================================
 GUI: HAUPTFENSTER, 3-SPALTEN HEADER & DASHBOARD (INKL. TOOL 12)
================================================================================
#>
function Start-AdminSuiteMainWindow {
    $mainForm = New-Object System.Windows.Forms.Form
    $mainForm.Text = Get-Text "Title"
    $mainForm.Size = New-Object System.Drawing.Size(980, 1030)
    $mainForm.StartPosition = "CenterScreen"
    $mainForm.FormBorderStyle = "FixedDialog"
    $mainForm.MaximizeBox = $false
    $mainForm.Font = New-Object System.Drawing.Font($script:UITheme.FontFamily, 9)
    $mainForm.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

    # --- HEADER PANEL (3 Spalten) ---
    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlHeader.Height = 230
    $pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(235, 242, 250)
    $mainForm.Controls.Add($pnlHeader)

    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.InitialDelay = 250
    $toolTip.ReshowDelay  = 100

    # CSV Export Button
    $btnExportCsv = New-Object System.Windows.Forms.Button
    $btnExportCsv.Location = New-Object System.Drawing.Point(720, 9)
    $btnExportCsv.Size = New-Object System.Drawing.Size(125, 26)
    $btnExportCsv.Text = "📥 CSV Export"
    $btnExportCsv.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 65)
    $btnExportCsv.ForeColor = [System.Drawing.Color]::White
    $btnExportCsv.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnExportCsv.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5, [System.Drawing.FontStyle]::Bold)
    $pnlHeader.Controls.Add($btnExportCsv)

    # Sprachumschalter (DE / EN)
    $btnLangEN = New-Object System.Windows.Forms.Button
    $btnLangEN.Location = New-Object System.Drawing.Point(860, 9)
    $btnLangEN.Size = New-Object System.Drawing.Size(42, 26)
    $btnLangEN.Text = "EN"
    $btnLangEN.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8, [System.Drawing.FontStyle]::Bold)
    $pnlHeader.Controls.Add($btnLangEN)

    $btnLangDE = New-Object System.Windows.Forms.Button
    $btnLangDE.Location = New-Object System.Drawing.Point(906, 9)
    $btnLangDE.Size = New-Object System.Drawing.Size(42, 26)
    $btnLangDE.Text = "DE"
    $btnLangDE.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8, [System.Drawing.FontStyle]::Bold)
    $pnlHeader.Controls.Add($btnLangDE)

    # Hauptüberschrift im Header
    $lblHeaderMain = New-Object System.Windows.Forms.Label
    $lblHeaderMain.Location = New-Object System.Drawing.Point(18, 10)
    $lblHeaderMain.Size = New-Object System.Drawing.Size(690, 24)
    $lblHeaderMain.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 11.5, [System.Drawing.FontStyle]::Bold)
    $lblHeaderMain.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $lblHeaderMain.Text = "💻 $global:localComputerName | $global:localUserName"
    $pnlHeader.Controls.Add($lblHeaderMain)

    # Trennlinie
    $lblHeaderLine = New-Object System.Windows.Forms.Label
    $lblHeaderLine.Location = New-Object System.Drawing.Point(18, 38)
    $lblHeaderLine.Size = New-Object System.Drawing.Size(935, 1)
    $lblHeaderLine.BackColor = [System.Drawing.Color]::FromArgb(203, 213, 225)
    $pnlHeader.Controls.Add($lblHeaderLine)

    # -------------------------------------------------------------
    # SPALTE 1: Betriebssystem & Domäne (Links)
    # -------------------------------------------------------------
    $lblCol1Title = New-Object System.Windows.Forms.Label
    $lblCol1Title.Location = New-Object System.Drawing.Point(18, 44)
    $lblCol1Title.Size = New-Object System.Drawing.Size(300, 18)
    $lblCol1Title.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $lblCol1Title.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $pnlHeader.Controls.Add($lblCol1Title)

    $lblCol1Content = New-Object System.Windows.Forms.Label
    $lblCol1Content.Location = New-Object System.Drawing.Point(18, 64)
    $lblCol1Content.Size = New-Object System.Drawing.Size(300, 155)
    $lblCol1Content.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $lblCol1Content.ForeColor = [System.Drawing.Color]::FromArgb(15, 23, 42)
    $pnlHeader.Controls.Add($lblCol1Content)

    # -------------------------------------------------------------
    # SPALTE 2: System & Hardware (Mitte - mit Klick-Kopierfunktion)
    # -------------------------------------------------------------
    $lblCol2Title = New-Object System.Windows.Forms.Label
    $lblCol2Title.Location = New-Object System.Drawing.Point(325, 44)
    $lblCol2Title.Size = New-Object System.Drawing.Size(300, 18)
    $lblCol2Title.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $lblCol2Title.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $pnlHeader.Controls.Add($lblCol2Title)

    # Computername (klickbar)
    $lblCompNameClick = New-Object System.Windows.Forms.Label
    $lblCompNameClick.Location = New-Object System.Drawing.Point(325, 64)
    $lblCompNameClick.Size = New-Object System.Drawing.Size(300, 18)
    $lblCompNameClick.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $lblCompNameClick.Cursor = [System.Windows.Forms.Cursors]::Hand
    $toolTip.SetToolTip($lblCompNameClick, "Klicken zum Kopieren des Computernamens ($global:localComputerName)")
    $lblCompNameClick.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($global:localComputerName)
        $toolTip.Show("✓ Computername '$global:localComputerName' kopiert!", $lblCompNameClick, 0, -22, 1200)
    })
    $pnlHeader.Controls.Add($lblCompNameClick)

    $lblHardwareMiddle = New-Object System.Windows.Forms.Label
    $lblHardwareMiddle.Location = New-Object System.Drawing.Point(325, 82)
    $lblHardwareMiddle.Size = New-Object System.Drawing.Size(300, 36)
    $lblHardwareMiddle.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $pnlHeader.Controls.Add($lblHardwareMiddle)

    # Seriennummer (klickbar)
    $lblSerialClick = New-Object System.Windows.Forms.Label
    $lblSerialClick.Location = New-Object System.Drawing.Point(325, 118)
    $lblSerialClick.Size = New-Object System.Drawing.Size(300, 18)
    $lblSerialClick.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $lblSerialClick.Cursor = [System.Windows.Forms.Cursors]::Hand
    $toolTip.SetToolTip($lblSerialClick, "Klicken zum Kopieren der Seriennummer ($global:localSerial)")
    $lblSerialClick.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($global:localSerial)
        $toolTip.Show("✓ Seriennummer '$global:localSerial' kopiert!", $lblSerialClick, 0, -22, 1200)
    })
    $pnlHeader.Controls.Add($lblSerialClick)

    $lblHardwareBottom = New-Object System.Windows.Forms.Label
    $lblHardwareBottom.Location = New-Object System.Drawing.Point(325, 136)
    $lblHardwareBottom.Size = New-Object System.Drawing.Size(300, 20)
    $lblHardwareBottom.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $pnlHeader.Controls.Add($lblHardwareBottom)

    # -------------------------------------------------------------
    # SPALTE 3: Entra ID / Cloud Status (Rechts - inkl. NgcSet & Status)
    # -------------------------------------------------------------
    $lblCol3Title = New-Object System.Windows.Forms.Label
    $lblCol3Title.Location = New-Object System.Drawing.Point(635, 44)
    $lblCol3Title.Size = New-Object System.Drawing.Size(320, 18)
    $lblCol3Title.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $lblCol3Title.ForeColor = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $pnlHeader.Controls.Add($lblCol3Title)

    $lblEntraTop = New-Object System.Windows.Forms.Label
    $lblEntraTop.Location = New-Object System.Drawing.Point(635, 64)
    $lblEntraTop.Size = New-Object System.Drawing.Size(320, 54)
    $lblEntraTop.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $pnlHeader.Controls.Add($lblEntraTop)

    # NgcSet (klickbar für Erklärung)
    $lblNgcSetClick = New-Object System.Windows.Forms.Label
    $lblNgcSetClick.Location = New-Object System.Drawing.Point(635, 118)
    $lblNgcSetClick.Size = New-Object System.Drawing.Size(320, 18)
    $lblNgcSetClick.Font = New-Object System.Drawing.Font("Consolas", 8.5, [System.Drawing.FontStyle]::Underline)
    $lblNgcSetClick.Cursor = [System.Windows.Forms.Cursors]::Hand
    $toolTip.SetToolTip($lblNgcSetClick, "Klicken für eine Erklärung zu NgcSet (Windows Hello for Business)")
    $lblNgcSetClick.Add_Click({
        [System.Windows.Forms.MessageBox]::Show(
@"
NgcSet (Next Generation Credential):
Aktueller Status: $global:localNgcSet

Bedeutung:
• YES: Für den angemeldeten Benutzer ist auf diesem Gerät ein Hardwareschlüssel für 'Windows Hello for Business' (WHfB) eingerichtet und an das TPM gekoppelt.
• Erlaubt die passwortlose Authentifizierung (PIN / Biometrie) an Microsoft Entra ID und Hybrid-Ressourcen.
• NO: Es ist kein WHfB-Container aktiv. Die Authentifizierung erfolgt klassisch über Kennwort / Kerberos.
"@,
            "Erklärung: NgcSet (Windows Hello for Business)",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })
    $pnlHeader.Controls.Add($lblNgcSetClick)

    $lblEntraBottom = New-Object System.Windows.Forms.Label
    $lblEntraBottom.Location = New-Object System.Drawing.Point(635, 136)
    $lblEntraBottom.Size = New-Object System.Drawing.Size(320, 60)
    $lblEntraBottom.Font = New-Object System.Drawing.Font("Consolas", 8.5)
    $pnlHeader.Controls.Add($lblEntraBottom)

    # --- CSV-EXPORT EVENT ---
    $btnExportCsv.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV-Datei (*.csv)|*.csv"
        $sfd.FileName = "SystemInfo_${global:localComputerName}_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv"

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $exportObj = [PSCustomObject]@{
                    "Computername"        = $global:localComputerName
                    "Benutzer"            = $global:localUserName
                    "OS Edition"          = $global:osCaption
                    "OS Build"            = $global:osBuildNumber
                    "OS Version"          = $global:osVersionDisplay
                    "Domaene"             = $global:localDomainName
                    "Logonserver"         = $global:localLogonServer
                    "Hersteller"          = $global:localManufacturer
                    "Modell"              = $global:localModel
                    "Seriennummer"        = $global:localSerial
                    "Systemtyp"           = $global:localSystemType
                    "Entra Join-Status"   = $global:localJoinStatus
                    "Azure Device Status" = $global:localAzureDevStat
                    "AzureAD PRT"         = $global:localAzureAdPrt
                    "NgcSet (WHfB)"       = $global:localNgcSet
                    "Tenant Name"         = $global:localTenantName
                    "Tenant ID"           = $global:localTenantId
                    "Device ID"           = $global:localDeviceId
                    "Export-Datum"        = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
                $exportObj | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                [System.Windows.Forms.MessageBox]::Show("Header-Informationen erfolgreich exportiert nach:`r`n$($sfd.FileName)", "Export erfolgreich", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim CSV-Export:`r`n$($_.Exception.Message)", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    # --- CLIENT TOOLS GROUPBOX ---
    $grpClient = New-Object System.Windows.Forms.GroupBox
    $grpClient.Location = New-Object System.Drawing.Point(18, 238)
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
    $grpAD.Location = New-Object System.Drawing.Point(18, 498)
    $grpAD.Size = New-Object System.Drawing.Size(935, 465)
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

    $btnTool12 = New-Object System.Windows.Forms.Button; $btnTool12.Location = "20, 386"; $btnTool12.Size = "895, 46"
    $btnTool12.Font = New-Object System.Drawing.Font($mainForm.Font.FontFamily, 8.5); $btnTool12.Add_Click({ Open-ToolUserPasswordAge }); $grpAD.Controls.Add($btnTool12)

    # --- UI REFRESH FUNKTION ---
    function Update-UI {
        $mainForm.Text  = Get-Text "Title"
        $grpClient.Text = Get-Text "CategoryClient"
        $grpAD.Text     = Get-Text "CategoryAD"

        $lblCol1Title.Text = Get-Text "LblHdrOS"
        $lblCol2Title.Text = if ($script:CurrentLang -eq "DE") { "SYSTEM & HARDWARE (Klick = Kopieren)" } else { "SYSTEM & HARDWARE (Click = Copy)" }
        $lblCol3Title.Text = Get-Text "LblHdrEntra"

        # Spalte 1: OS & Domäne
        $lblCol1Content.Text = @"
$("{0,-12}: {1}" -f (Get-Text "LblOS"), $global:osCaption)
$("{0,-12}: {1}" -f (Get-Text "LblBuild"), $global:osBuildNumber)
$("{0,-12}: {1}" -f (Get-Text "LblVersion"), $global:osVersionDisplay)
$("{0,-12}: {1}" -f (Get-Text "LblDomain"), $global:localDomainName)
$("{0,-12}: {1}" -f (Get-Text "LblLogonServer"), $global:localLogonServer)
"@

        # Spalte 2: System & Hardware
        $lblCompNameClick.Text = "{0,-13}: {1}" -f (Get-Text "LblCompName"), $global:localComputerName
        $lblHardwareMiddle.Text = @"
$("{0,-13}: {1}" -f (Get-Text "LblManuf"), $global:localManufacturer)
$("{0,-13}: {1}" -f (Get-Text "LblModel"), $global:localModel)
"@
        $lblSerialClick.Text    = "{0,-13}: {1}" -f (Get-Text "LblSerial"), $global:localSerial
        $lblHardwareBottom.Text = "{0,-13}: {1}" -f (Get-Text "LblSysType"), $global:localSystemType

        # Spalte 3: Entra ID / Cloud Status
        $lblEntraTop.Text = @"
$("{0,-12}: {1}" -f (Get-Text "LblJoinStatus"), $global:localJoinStatus)
$("{0,-12}: {1}" -f "DeviceStatus", $global:localAzureDevStat)
$("{0,-12}: {1}" -f (Get-Text "LblPrtStatus"), $global:localAzureAdPrt)
"@
        $lblNgcSetClick.Text = "{0,-12}: {1} 🛈 (Info)" -f "NgcSet", $global:localNgcSet
        $lblNgcSetClick.ForeColor = if ($global:localNgcSet -eq "YES") { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::FromArgb(15, 23, 42) }

        $lblEntraBottom.Text = @"
$("{0,-12}: {1}" -f (Get-Text "LblTenantName"), $global:localTenantName)
$("{0,-12}: {1}" -f (Get-Text "LblTenantId"), $(if ($global:localTenantId.Length -gt 16) { $global:localTenantId.Substring(0,13) + "..." } else { $global:localTenantId }))
$("{0,-12}: {1}" -f (Get-Text "LblDeviceId"), $(if ($global:localDeviceId.Length -gt 16) { $global:localDeviceId.Substring(0,13) + "..." } else { $global:localDeviceId }))
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
        $btnTool12.Text = Get-Text "BtnTool12"

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