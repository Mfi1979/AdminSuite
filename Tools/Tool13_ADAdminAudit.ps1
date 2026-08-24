Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

# Hauptfenster
$form = New-Object System.Windows.Forms.Form
$form.Text = "Active Directory Tier-0 & Identity Security Audit Tool"
$form.Size = New-Object System.Drawing.Size(1380, 820)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"

# =========================================================================
# TAB 1: Admin- & Notfall-Konten (Volle Breite)
# =========================================================================
$tabAdmins = New-Object System.Windows.Forms.TabPage
$tabAdmins.Text = "1. Admin- & Notfall-Konten"

$btnScanAdmins = New-Object System.Windows.Forms.Button
$btnScanAdmins.Text = "Admin-Konten scannen"
$btnScanAdmins.Location = New-Object System.Drawing.Point(15, 12)
$btnScanAdmins.Size = New-Object System.Drawing.Size(180, 30)

$lblLegendAdm = New-Object System.Windows.Forms.Label
$lblLegendAdm.Text = "Legende:  [Grün] OK / Gehärtet  |  [Gelb/Orange] Abweichung (PW > 365d, Protected Users fehlt)  |  [Rot] Deaktiviert"
$lblLegendAdm.Location = New-Object System.Drawing.Point(210, 18)
$lblLegendAdm.AutoSize = $true
$lblLegendAdm.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

# 1. Accounts-Tabelle (Volle Breite)
$lblAdmTable = New-Object System.Windows.Forms.Label
$lblAdmTable.Text = "1. Admin- oder Notfall-Konto auswählen:"
$lblAdmTable.Location = New-Object System.Drawing.Point(15, 48)
$lblAdmTable.AutoSize = $true
$lblAdmTable.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$gridAdmins = New-Object System.Windows.Forms.DataGridView
$gridAdmins.Location = New-Object System.Drawing.Point(15, 70)
$gridAdmins.Size = New-Object System.Drawing.Size(1325, 270)
$gridAdmins.ReadOnly = $true
$gridAdmins.AutoSizeColumnsMode = "AllCells"
$gridAdmins.AllowUserToAddRows = $false
$gridAdmins.SelectionMode = "FullRowSelect"
$gridAdmins.MultiSelect = $false

# 2. Detail-Gruppen (Volle Breite)
$lblAdmGroups = New-Object System.Windows.Forms.Label
$lblAdmGroups.Text = "2. Gruppenmitgliedschaften des Kontos:"
$lblAdmGroups.Location = New-Object System.Drawing.Point(15, 350)
$lblAdmGroups.AutoSize = $true
$lblAdmGroups.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$gridAdmGroups = New-Object System.Windows.Forms.DataGridView
$gridAdmGroups.Location = New-Object System.Drawing.Point(15, 375)
$gridAdmGroups.Size = New-Object System.Drawing.Size(1325, 355)
$gridAdmGroups.ReadOnly = $true
$gridAdmGroups.AutoSizeColumnsMode = "Fill"
$gridAdmGroups.AllowUserToAddRows = $false

$tabAdmins.Controls.Add($btnScanAdmins)
$tabAdmins.Controls.Add($lblLegendAdm)
$tabAdmins.Controls.Add($lblAdmTable)
$tabAdmins.Controls.Add($gridAdmins)
$tabAdmins.Controls.Add($lblAdmGroups)
$tabAdmins.Controls.Add($gridAdmGroups)

# =========================================================================
# TAB 2: Spezielle Dienst- & Systemkonten (krbtgt, Entra Sync, Seamless SSO)
# =========================================================================
$tabServices = New-Object System.Windows.Forms.TabPage
$tabServices.Text = "2. Dienst- & Systemkonten (krbtgt / Entra ID)"

$btnScanServices = New-Object System.Windows.Forms.Button
$btnScanServices.Text = "Dienstkonten neu scannen"
$btnScanServices.Location = New-Object System.Drawing.Point(15, 12)
$btnScanServices.Size = New-Object System.Drawing.Size(190, 30)

$lblLegendSvc = New-Object System.Windows.Forms.Label
$lblLegendSvc.Text = "Legende:  [Grün] OK / Best-Practice  |  [Gelb/Orange] Schlüsselalter/Warnung  |  [Rot] Kritische Fehlkonfiguration"
$lblLegendSvc.Location = New-Object System.Drawing.Point(220, 18)
$lblLegendSvc.AutoSize = $true
$lblLegendSvc.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

# 1. Accounts-Tabelle
$lblSvcTable = New-Object System.Windows.Forms.Label
$lblSvcTable.Text = "1. Dienst- / System-Konto auswählen:"
$lblSvcTable.Location = New-Object System.Drawing.Point(15, 48)
$lblSvcTable.AutoSize = $true
$lblSvcTable.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$gridServices = New-Object System.Windows.Forms.DataGridView
$gridServices.Location = New-Object System.Drawing.Point(15, 70)
$gridServices.Size = New-Object System.Drawing.Size(800, 270)
$gridServices.ReadOnly = $true
$gridServices.AutoSizeColumnsMode = "AllCells"
$gridServices.AllowUserToAddRows = $false
$gridServices.SelectionMode = "FullRowSelect"
$gridServices.MultiSelect = $false

# 2. Detail-Gruppen
$lblSvcGroups = New-Object System.Windows.Forms.Label
$lblSvcGroups.Text = "2. Gruppenmitgliedschaften des Dienstkontos:"
$lblSvcGroups.Location = New-Object System.Drawing.Point(15, 350)
$lblSvcGroups.AutoSize = $true
$lblSvcGroups.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$gridSvcGroups = New-Object System.Windows.Forms.DataGridView
$gridSvcGroups.Location = New-Object System.Drawing.Point(15, 375)
$gridSvcGroups.Size = New-Object System.Drawing.Size(800, 355)
$gridSvcGroups.ReadOnly = $true
$gridSvcGroups.AutoSizeColumnsMode = "Fill"
$gridSvcGroups.AllowUserToAddRows = $false

# 3. Erklärung Rechts
$lblSvcDesc = New-Object System.Windows.Forms.Label
$lblSvcDesc.Text = "3. Beschreibung, Funktion & Sicherheitsleitfaden:"
$lblSvcDesc.Location = New-Object System.Drawing.Point(830, 48)
$lblSvcDesc.AutoSize = $true
$lblSvcDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$txtSvcDesc = New-Object System.Windows.Forms.TextBox
$txtSvcDesc.Location = New-Object System.Drawing.Point(830, 70)
$txtSvcDesc.Size = New-Object System.Drawing.Size(515, 660)
$txtSvcDesc.Multiline = $true
$txtSvcDesc.ReadOnly = $true
$txtSvcDesc.ScrollBars = "Vertical"
$txtSvcDesc.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
$txtSvcDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

$tabServices.Controls.Add($btnScanServices)
$tabServices.Controls.Add($lblLegendSvc)
$tabServices.Controls.Add($lblSvcTable)
$tabServices.Controls.Add($gridServices)
$tabServices.Controls.Add($lblSvcGroups)
$tabServices.Controls.Add($gridSvcGroups)
$tabServices.Controls.Add($lblSvcDesc)
$tabServices.Controls.Add($txtSvcDesc)

# =========================================================================
# TAB 3: Privilegierte Gruppen & Mitglieder
# =========================================================================
$tabGroups = New-Object System.Windows.Forms.TabPage
$tabGroups.Text = "3. Privilegierte Gruppen & Mitglieder"

$btnScanGroups = New-Object System.Windows.Forms.Button
$btnScanGroups.Text = "Gruppen neu laden"
$btnScanGroups.Location = New-Object System.Drawing.Point(15, 12)
$btnScanGroups.Size = New-Object System.Drawing.Size(160, 30)

$lblLegendGrp = New-Object System.Windows.Forms.Label
$lblLegendGrp.Text = "Legende:  [Rot] Sicherheitsrisiko (z.B. Schema-Admins nicht leer)  |  [Grün] Empfohlener Zustand"
$lblLegendGrp.Location = New-Object System.Drawing.Point(190, 18)
$lblLegendGrp.AutoSize = $true
$lblLegendGrp.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)

# 1. Gruppen-Tabelle
$lblGroups = New-Object System.Windows.Forms.Label
$lblGroups.Text = "1. Gruppe auswählen:"
$lblGroups.Location = New-Object System.Drawing.Point(15, 48)
$lblGroups.AutoSize = $true
$lblGroups.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$gridGroups = New-Object System.Windows.Forms.DataGridView
$gridGroups.Location = New-Object System.Drawing.Point(15, 70)
$gridGroups.Size = New-Object System.Drawing.Size(800, 270)
$gridGroups.ReadOnly = $true
$gridGroups.SelectionMode = "FullRowSelect"
$gridGroups.MultiSelect = $false
$gridGroups.AutoSizeColumnsMode = "AllCells"
$gridGroups.AllowUserToAddRows = $false

# 2. Gruppen-Mitglieder
$lblMembers = New-Object System.Windows.Forms.Label
$lblMembers.Text = "2. Mitglieder der ausgewählten Gruppe:"
$lblMembers.Location = New-Object System.Drawing.Point(15, 350)
$lblMembers.AutoSize = $true
$lblMembers.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$gridMembers = New-Object System.Windows.Forms.DataGridView
$gridMembers.Location = New-Object System.Drawing.Point(15, 375)
$gridMembers.Size = New-Object System.Drawing.Size(800, 355)
$gridMembers.ReadOnly = $true
$gridMembers.AutoSizeColumnsMode = "Fill"
$gridMembers.AllowUserToAddRows = $false

# 3. Gruppen-Erklärung
$lblDesc = New-Object System.Windows.Forms.Label
$lblDesc.Text = "3. Beschreibung & Sicherheits-Erklärung:"
$lblDesc.Location = New-Object System.Drawing.Point(830, 48)
$lblDesc.AutoSize = $true
$lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$txtDescription = New-Object System.Windows.Forms.TextBox
$txtDescription.Location = New-Object System.Drawing.Point(830, 70)
$txtDescription.Size = New-Object System.Drawing.Size(515, 660)
$txtDescription.Multiline = $true
$txtDescription.ReadOnly = $true
$txtDescription.ScrollBars = "Vertical"
$txtDescription.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 252)
$txtDescription.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

$tabGroups.Controls.Add($btnScanGroups)
$tabGroups.Controls.Add($lblLegendGrp)
$tabGroups.Controls.Add($lblGroups)
$tabGroups.Controls.Add($gridGroups)
$tabGroups.Controls.Add($lblMembers)
$tabGroups.Controls.Add($gridMembers)
$tabGroups.Controls.Add($lblDesc)
$tabGroups.Controls.Add($txtDescription)

$tabControl.TabPages.Add($tabAdmins)
$tabControl.TabPages.Add($tabServices)
$tabControl.TabPages.Add($tabGroups)
$form.Controls.Add($tabControl)

# =========================================================================
# SIDs & Hilfsfunktionen
# =========================================================================
$domainSid = (Get-ADDomain).DomainSID.Value

$privilegedGroupIdentifiers = @(
    "$domainSid-518",  # Schema Admins
    "$domainSid-519",  # Enterprise Admins
    "$domainSid-512",  # Domain Admins
    "S-1-5-32-544",    # Builtin Administrators
    "$domainSid-525",  # Protected Users
    "$domainSid-520",  # Group Policy Creator Owners
    "ADS-FGPP-DADM",   # Custom Gruppen
    "GRP-FGPP-DAdm",
    "DnsAdmins",
    "DnsUpdateProxy",
    "ADSyncAdmins",
    "KLAdmins"
)

function Get-ResolvedADGroup {
    param($id)
    try {
        return (Get-ADGroup -Identity $id -Properties Members, MemberOf, Description, whenCreated -ErrorAction Stop)
    } catch {
        return $null
    }
}

function Get-GroupOriginType {
    param($grpSid, $grpCreated, $refDate)
    $isDefault = $false
    if ($grpSid.StartsWith("S-1-5-32-")) {
        $isDefault = $true
    } elseif ($grpSid -match '-(\d+)$') {
        $rid = [int]$matches[1]
        if ($rid -lt 1000) { $isDefault = $true }
    }
    if ($refDate -and $grpCreated) {
        if ((Get-Date $grpCreated).Date -eq $refDate) {
            $isDefault = $true
        }
    }
    return $(if ($isDefault) { "Default (System)" } else { "Manuell (Custom)" })
}

# Wissensdatenbank: Dienstkonten
function Get-ServiceAccountKnowledgeBaseText {
    param($userName, $role, $status, $pwdDateStr, $pwdAgeDays, $isProtected, $isNotDelegated, $findings, $dn)

    $text = "DIENSTKONTO: $userName`r`n"
    $text += "FUNKTION: $role`r`n"
    $text += "STATUS: $status`r`n"
    $text += "LETZTE PW-ÄNDERUNG: $pwdDateStr ($pwdAgeDays Tage)`r`n"
    $text += "PROTECTED USERS: $isProtected | NICHT DELEGIERBAR: $isNotDelegated`r`n"
    $text += "DN: $dn`r`n"
    $text += "====================================================`r`n`r`n"

    if ($userName -match "^krbtgt") {
        $text += "[FUNKTION & HINTERGRUND]`r`n"
        $text += "Das Kerberos Key Distribution Center (KDC) Masterkonto. Es signiert und verschlüsselt alle Ticket Granting Tickets (TGTs) im Active Directory.`r`n`r`n"
        $text += "[SICHERHEITSBEWERTUNG: TIER 0 / HÖCHSTE KRITIKALITÄT]`r`n"
        $text += "Erlangt ein Angreifer das Kennwort/den NTLM-Hash von 'krbtgt', kann er 'Golden Tickets' schmieden und besitzt uneingeschränkten Vollzugriff.`r`n`r`n"
        $text += "[BEST PRACTICE EMPFEHLUNGEN]`r`n"
        $text += "• Regelmäßiger Double-Reset des Kennworts (z. B. alle 180 Tage) via PowerShell-Skript 'Reset-KrbTgtKey.ps1'.`r`n"
        $text += "• Das Konto MUSS deaktiviert bleiben (Windows-Standard).`r`n"
        $text += "• krbtgt darf NICHT in 'Protected Users' aufgenommen werden!"
    }
    elseif ($userName -match "^MSOL_" -or $userName -match "^Sync_" -or $userName -match "^AAD_") {
        $text += "[FUNKTION & HINTERGRUND]`r`n"
        $text += "Dienstkonto von Microsoft Entra Connect (Azure AD Sync). Synchronisiert Benutzer, Attribute und Kennworthashes (PHS) mit Microsoft Entra ID.`r`n`r`n"
        $text += "[SICHERHEITSBEWERTUNG: TIER 0 DIENSTKONTO]`r`n"
        $text += "Besitzt 'Replicating Directory Changes' (DCSync-Rechte), um Passwort-Hashes direkt aus der NTDS.dit-Datenbank zu lesen.`r`n`r`n"
        $text += "[BEST PRACTICE EMPFEHLUNGEN]`r`n"
        $text += "• ACHTUNG: Darf NICHT in der Gruppe 'Protected Users' sein, da das NTLM/RPC-Verbot die Synchronisation sofort abbricht!`r`n"
        $text += "• 'Konto ist vertraulich und kann nicht delegiert werden' aktivieren.`r`n"
        $text += "• Keine administrativen Gruppen wie 'Domänen-Admins' zuweisen (DCSync-Rechte reichen vollständig aus)."
    }
    elseif ($userName -match "^AZUREADSSOACC") {
        $text += "[FUNKTION & HINTERGRUND]`r`n"
        $text += "Kerberos-Computerkonto / Secret Holder für Microsoft Entra Seamless Single Sign-On (Nahtloses SSO).`r`n`r`n"
        $text += "[BEST PRACTICE EMPFEHLUNGEN]`r`n"
        $text += "• Mindestens alle 30-90 Tage einen Key-Rollover mittels PowerShell (Update-AzureADSSOForest) ausführen, um veraltete Schlüssel zu erneuern."
    }
    elseif ($role -match "gMSA") {
        $text += "[FUNKTION & HINTERGRUND]`r`n"
        $text += "Group Managed Service Account (gMSA). Modernes, vom Active Directory automatisch rotiertes Dienstkonto (128-Bit Kennwort alle 30 Tage).`r`n`r`n"
        $text += "[BEST PRACTICE EMPFEHLUNGEN]`r`n"
        $text += "• Höchster empfohlener Sicherheitsstandard für Windows-Dienste und geplante Aufgaben.`r`n"
        $text += "• Keine manuelle Passwortverwaltung nötig."
    }
    else {
        $text += "[FUNKTION & HINTERGRUND]`r`n"
        $text += "Dienstkonto mit registriertem Service Principal Name (SPN) oder Kerberos-Dienstbindung.`r`n`r`n"
        $text += "[SICHERHEITSHINWEIS: KERBEROASTING]`r`n"
        $text += "Konten mit gesetztem SPN können Ziel von Offline-Kerberoasting-Angriffen sein.`r`n"
        $text += "• Mindestens 25+ Zeichen langes Passwort sicherstellen.`r`n"
        $text += "• Falls möglich, auf gMSA migrieren."
    }

    return $text
}

# Wissensdatenbank: Gruppen
function Get-GroupKnowledgeBaseText {
    param($grpName, $sid, $adDescription, $createdDate, $originType)

    $text = "GRUPPE: $grpName`r`n"
    $text += "HERKUNFT: $originType`r`n"
    $text += "ERSTELLT AM: $createdDate`r`n"
    $text += "SID: $sid`r`n"
    if ($adDescription) { $text += "AD-Beschreibung: $adDescription`r`n" }
    $text += "====================================================`r`n`r`n"

    if ($sid -match "-518$") {
        $text += "[FUNKTION & AUFGABE]`r`n"
        $text += "Verwaltet das Active Directory Schema. Mitglieder können Objektklassen und Attribute der gesamten Gesamtstruktur erweitern.`r`n`r`n"
        $text += "[SICHERHEITSBEWERTUNG: KRITISCH (Tier 0)]`r`n"
        $text += "• Diese Gruppe MUSS im normalen Alltagsbetrieb LEER sein!`r`n"
        $text += "• Konten werden ausschließlich temporär (Just-in-Time) für Schema-Updates hinzugefügt und direkt wieder entfernt."
    }
    elseif ($sid -match "-519$") {
        $text += "[FUNKTION & AUFGABE]`r`n"
        $text += "Administrative Vollzugriffsrechte über die gesamte Gesamtstruktur (Forest Root).`r`n`r`n"
        $text += "[BEST PRACTICE EMPFEHLUNG]`r`n"
        $text += "• Im Alltagsbetrieb LEER halten. Nur das primäre Notfall-Konto oder temporäre Admins für Gesamtstruktur-Wartungen zulassen."
    }
    elseif ($sid -match "-512$") {
        $text += "[FUNKTION & AUFGABE]`r`n"
        $text += "Vollzugriff auf alle Objekte, Richtlinien und Domain Controller innerhalb dieser Domäne.`r`n`r`n"
        $text += "[BEST PRACTICE EMPFEHLUNG]`r`n"
        $text += "• Anzahl der Konten so gering wie möglich halten.`r`n"
        $text += "• Persönliche Domain Admins MÜSSEN in die Gruppe 'Protected Users' aufgenommen werden."
    }
    elseif ($sid -match "-525$") {
        $text += "[FUNKTION & AUFGABE]`r`n"
        $text += "Erzwingt strenge Protokollbeschränkungen auf Kerberos- und Betriebssystemebene:`r`n"
        $text += "• Verhindert NTLM, Digest und CredSSP komplett.`r`n"
        $text += "• Erzwingt AES-Verschlüsselung und verhindert Credential Caching.`r`n"
        $text += "• Begrenzt TGT-Lebensdauer auf 4 Stunden.`r`n`r`n"
        $text += "[BEST PRACTICE EMPFEHLUNG]`r`n"
        $text += "• Alle regulären, menschlichen Admins hier hinein.`r`n"
        $text += "• VORSICHT: Notfall-Konten (Glass-Break) und Dienstkonten (krbtgt, MSOL_*) dürfen NICHT hier hinein!"
    }
    elseif ($grpName -eq "DnsAdmins") {
        $text += "[FUNKTION & AUFGABE]`r`n"
        $text += "Verwaltet den Microsoft DNS-Dienst.`r`n`r`n"
        $text += "[SICHERHEITSHINWEIS]`r`n"
        $text += "DnsAdmins können über DLL-Injektion (ServerLevelPluginDll) SYSTEM-Rechte auf Domain Controllern erlangen. Nur absolut vertrauenswürdigen Tier-0-Konten zuweisen."
    }
    elseif ($grpName -eq "DnsUpdateProxy") {
        $text += "[FUNKTION & AUFGABE]`r`n"
        $text += "Spezialgruppe für DHCP-Server zur dynamischen Registrierung älterer Clients.`r`n`r`n"
        $text += "[SICHERHEITSHINWEIS: GEFAHR]`r`n"
        $text += "Hier dürfen NIEMALS Benutzer- oder Admin-Konten enthalten sein!"
    }
    else {
        $text += "[ALLGEMEINE INFORMATION]`r`n"
        $text += "Privilegierte Sicherheitsgruppe im Active Directory."
    }

    return $text
}

# =========================================================================
# DataBindingComplete Farblogik
# =========================================================================
$gridAdmins.Add_DataBindingComplete({
    foreach ($row in $gridAdmins.Rows) {
        $status = [string]$row.Cells["Status (Aktiv)"].Value
        $befund = [string]$row.Cells["Empfehlung / Audit-Befund"].Value

        if ($status -ne "Aktiv" -and $status -ne "") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::Salmon
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
        }
        elseif ($befund -match "WARNUNG" -or $befund -match "PW > 365" -or $befund -match "älter als 365" -or $befund -match "Härtung") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LemonChiffon
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkOrange
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::Khaki
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
        }
        else {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 235)
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 235, 200)
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
        }
    }
})

$gridServices.Add_DataBindingComplete({
    foreach ($row in $gridServices.Rows) {
        $befund = [string]$row.Cells["Empfehlung / Audit-Befund"].Value

        if ($befund -match "FEHLER" -or $befund -match "KRITISCH") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::Salmon
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
        }
        elseif ($befund -match "Warnung" -or $befund -match "älter als" -or $befund -match "empfohlen" -or $befund -match "setzen") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LemonChiffon
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkOrange
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::Khaki
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
        }
        else {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 235)
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 235, 200)
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
        }
    }
})

$gridGroups.Add_DataBindingComplete({
    foreach ($row in $gridGroups.Rows) {
        $bewertung = [string]$row.Cells["Sicherheits-Bewertung"].Value
        if ($bewertung -match "CRITICAL") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkRed
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::Salmon
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
            $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($gridGroups.Font, [System.Drawing.FontStyle]::Bold)
        }
        elseif ($bewertung -match "Hinweis") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LemonChiffon
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkOrange
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::Khaki
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
        }
        elseif ($bewertung -match "Optimal") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235, 247, 235)
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGreen
            $row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(200, 235, 200)
            $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
        }
    }
})

# =========================================================================
# Event: Tab 1 - Admin-Konten Scan
# =========================================================================
$btnScanAdmins.Add_Click({
    $gridAdmins.DataSource = $null
    $table = New-Object System.Data.DataTable
    
    [void]$table.Columns.Add("Klassifizierung")
    [void]$table.Columns.Add("SamAccountName")
    [void]$table.Columns.Add("Name")
    [void]$table.Columns.Add("Status (Aktiv)")
    [void]$table.Columns.Add("PW zuletzt geändert")
    [void]$table.Columns.Add("PW-Alter (Tage)", [int])
    [void]$table.Columns.Add("Protected Users?")
    [void]$table.Columns.Add("Nicht Delegierbar")
    [void]$table.Columns.Add("Anzahl Gruppen", [int])
    [void]$table.Columns.Add("Empfehlung / Audit-Befund")
    [void]$table.Columns.Add("DistinguishedName")

    $discoveredUsers = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($id in $privilegedGroupIdentifiers) {
        $grp = Get-ResolvedADGroup $id
        if ($grp) {
            try {
                $members = Get-ADGroupMember -Identity $grp.DistinguishedName -Recursive -ErrorAction SilentlyContinue | Where-Object { $_.objectClass -eq "user" }
                foreach ($m in $members) { [void]$discoveredUsers.Add($m.SamAccountName) }
            } catch {}
        }
    }

    try {
        Get-ADUser -Filter {adminCount -eq 1} -ErrorAction SilentlyContinue | ForEach-Object {
            [void]$discoveredUsers.Add($_.SamAccountName)
        }
    } catch {}

    foreach ($userName in ($discoveredUsers | Sort-Object)) {
        if ($userName -match "^krbtgt" -or $userName -match "^MSOL_" -or $userName -match "^Sync_" -or $userName -match "^AAD_" -or $userName -match "^AZUREADSSOACC") {
            continue
        }

        try {
            $user = Get-ADUser -Identity $userName -Properties MemberOf, AccountNotDelegated, Enabled, SID, PasswordLastSet, PasswordNeverExpires -ErrorAction Stop
            $groups = $user.MemberOf | ForEach-Object { (Get-ADGroup -Identity $_).Name } | Sort-Object
            $isProtected = if ($groups -contains "Protected Users") { "JA" } else { "NEIN" }
            $isNotDelegated = if ($user.AccountNotDelegated) { "JA" } else { "NEIN" }
            $isRid500 = $user.SID.Value.EndsWith("-500")

            $pwdLastSet = $user.PasswordLastSet
            $pwdDateStr = "-"
            $pwdAgeDays = 0
            $isPwdTooOld = $false

            if ($pwdLastSet) {
                $pwdDateStr = (Get-Date $pwdLastSet).ToString("dd.MM.yyyy HH:mm")
                $pwdAgeDays = [int][math]::Round(((Get-Date) - $pwdLastSet).TotalDays, 0)
                if ($pwdAgeDays -gt 365) { $isPwdTooOld = $true }
            } else {
                $pwdDateStr = "Nie / Unbekannt"
                $pwdAgeDays = 9999
                $isPwdTooOld = $true
            }

            $role = "Standard Admin (Tier-0/1)"
            $findings = [System.Collections.Generic.List[string]]::new()

            if ($isRid500 -or $userName -match "exit|break|emergency|notfall|adm-adex") {
                $role = "Glass-Break / Notfall-Admin"
                if ($isProtected -eq "JA") {
                    $findings.Add("WARNUNG: Glass-Break darf NICHT in Protected Users sein!")
                }
                if ($isNotDelegated -ne "JA") {
                    $findings.Add("Flag 'Nicht delegierbar' setzen")
                }
            } else {
                if ($groups -contains "Domänen-Admins" -or $groups -contains "Domain Admins") {
                    $role = "Regulärer Domain Admin (Tier-0)"
                    if ($isProtected -ne "JA") {
                        $findings.Add("Härtung: Zu 'Protected Users' hinzufügen")
                    }
                } else {
                    $role = "Delegierter / Server-Admin"
                }
            }

            if ($isPwdTooOld) {
                if ($pwdAgeDays -eq 9999) {
                    $findings.Add("⚠️ PW wurde NIE geändert!")
                } else {
                    $findings.Add("⚠️ PW > 365 Tage alt ($pwdAgeDays d)")
                }
            }

            if (-not $user.Enabled) {
                $findings.Add("Konto ist DEAKTIVIERT")
            }

            $finalFinding = if ($findings.Count -gt 0) { $findings -join " | " } else { "OK" }

            [void]$table.Rows.Add(
                $role,
                $user.SamAccountName,
                $user.Name,
                $(if($user.Enabled){"Aktiv"}else{"Gesperrt/Deaktiviert"}),
                $pwdDateStr,
                $pwdAgeDays,
                $isProtected,
                $isNotDelegated,
                @($groups).Count,
                $finalFinding,
                $user.DistinguishedName
            )
        } catch {}
    }

    $gridAdmins.DataSource = $table
    $gridAdmins.Columns["DistinguishedName"].Visible = $false
})

$gridAdmins.Add_SelectionChanged({
    if ($gridAdmins.SelectedRows.Count -gt 0) {
        $row = $gridAdmins.SelectedRows[0]
        $selectedUser = [string]$row.Cells["SamAccountName"].Value
        $selectedDn = [string]$row.Cells["DistinguishedName"].Value

        $lblAdmGroups.Text = "2. Gruppenmitgliedschaften des Kontos: [$selectedUser]"

        $tableUserGroups = New-Object System.Data.DataTable
        [void]$tableUserGroups.Columns.Add("Gruppen-Name")
        [void]$tableUserGroups.Columns.Add("Herkunft / Typ")
        [void]$tableUserGroups.Columns.Add("Erstellungsdatum")
        [void]$tableUserGroups.Columns.Add("SID")
        [void]$tableUserGroups.Columns.Add("Beschreibung")

        $domainAdminsGroup = Get-ResolvedADGroup "$domainSid-512"
        $daCreationDate = if ($domainAdminsGroup.whenCreated) { (Get-Date $domainAdminsGroup.whenCreated).Date } else { $null }

        try {
            $userObj = Get-ADUser -Identity $selectedDn -Properties memberOf -ErrorAction Stop
            $allGroupDns = [System.Collections.Generic.List[string]]::new()
            if ($userObj.memberOf) {
                foreach ($g in $userObj.memberOf) { $allGroupDns.Add($g) }
            }

            if ($allGroupDns.Count -eq 0) {
                [void]$tableUserGroups.Rows.Add("-- Keine direkten Gruppen --", "-", "-", "-", "-")
            } else {
                foreach ($grpDn in ($allGroupDns | Sort-Object)) {
                    try {
                        $grpObj = Get-ADGroup -Identity $grpDn -Properties Description, whenCreated, SID -ErrorAction Stop
                        $origin = Get-GroupOriginType -grpSid $grpObj.SID.Value -grpCreated $grpObj.whenCreated -refDate $daCreationDate
                        $createdStr = if ($grpObj.whenCreated) { (Get-Date $grpObj.whenCreated).ToString("dd.MM.yyyy HH:mm") } else { "-" }

                        [void]$tableUserGroups.Rows.Add(
                            $grpObj.Name,
                            $origin,
                            $createdStr,
                            $grpObj.SID.Value,
                            $grpObj.Description
                        )
                    } catch {
                        [void]$tableUserGroups.Rows.Add($grpDn, "Unbekannt", "-", "-", "Gruppe konnte nicht aufgelöst werden")
                    }
                }
            }
        } catch {
            [void]$tableUserGroups.Rows.Add("Fehler", "Gruppen konnten nicht geladen werden", "-", "-", "-")
        }

        $gridAdmGroups.DataSource = $tableUserGroups
    }
})

# =========================================================================
# Event: TAB 2 - Dienstkonten Scan (krbtgt, MSOL, SSO, gMSA)
# =========================================================================
$btnScanServices.Add_Click({
    $gridServices.DataSource = $null
    $table = New-Object System.Data.DataTable
    
    [void]$table.Columns.Add("Dienst-Funktion")
    [void]$table.Columns.Add("SamAccountName")
    [void]$table.Columns.Add("Name")
    [void]$table.Columns.Add("Status (Aktiv)")
    [void]$table.Columns.Add("Schlüssel/PW geändert")
    [void]$table.Columns.Add("Alter (Tage)", [int])
    [void]$table.Columns.Add("Protected Users?")
    [void]$table.Columns.Add("Nicht Delegierbar")
    [void]$table.Columns.Add("Empfehlung / Audit-Befund")
    [void]$table.Columns.Add("DistinguishedName")

    # 1. krbtgt Konten
    try {
        $krbtgtAccounts = Get-ADUser -Filter {SamAccountName -like "krbtgt*"} -Properties MemberOf, AccountNotDelegated, Enabled, SID, PasswordLastSet, Description -ErrorAction SilentlyContinue
        foreach ($user in $krbtgtAccounts) {
            $groups = $user.MemberOf | ForEach-Object { (Get-ADGroup -Identity $_).Name } | Sort-Object
            $isProtected = if ($groups -contains "Protected Users") { "JA" } else { "NEIN" }
            $isNotDelegated = if ($user.AccountNotDelegated) { "JA" } else { "NEIN" }

            $pwdLastSet = $user.PasswordLastSet
            $pwdDateStr = if ($pwdLastSet) { (Get-Date $pwdLastSet).ToString("dd.MM.yyyy HH:mm") } else { "Nie" }
            $pwdAgeDays = if ($pwdLastSet) { [int][math]::Round(((Get-Date) - $pwdLastSet).TotalDays, 0) } else { 9999 }

            $findings = [System.Collections.Generic.List[string]]::new()
            if ($pwdAgeDays -gt 180) { $findings.Add("⚠️ krbtgt-PW älter als 180 Tage ($pwdAgeDays d) -> Reset empfohlen!") }
            if ($isProtected -eq "JA") { $findings.Add("FEHLER: krbtgt darf NICHT in Protected Users sein!") }
            if ($user.Enabled) { $findings.Add("Hinweis: krbtgt sollte deaktiviert sein") }
            $finalFinding = if ($findings.Count -gt 0) { $findings -join " | " } else { "OK" }

            [void]$table.Rows.Add("Kerberos Master Key (krbtgt)", $user.SamAccountName, $user.Name, $(if($user.Enabled){"Aktiv"}else{"Deaktiviert"}), $pwdDateStr, $pwdAgeDays, $isProtected, $isNotDelegated, $finalFinding, $user.DistinguishedName)
        }
    } catch {}

    # 2. Entra Connect Sync Konten
    try {
        $syncAccounts = Get-ADUser -Filter {SamAccountName -like "MSOL_*" -or SamAccountName -like "Sync_*" -or SamAccountName -like "AAD_*"} -Properties MemberOf, AccountNotDelegated, Enabled, SID, PasswordLastSet, Description -ErrorAction SilentlyContinue
        foreach ($user in $syncAccounts) {
            $groups = $user.MemberOf | ForEach-Object { (Get-ADGroup -Identity $_).Name } | Sort-Object
            $isProtected = if ($groups -contains "Protected Users") { "JA" } else { "NEIN" }
            $isNotDelegated = if ($user.AccountNotDelegated) { "JA" } else { "NEIN" }

            $pwdLastSet = $user.PasswordLastSet
            $pwdDateStr = if ($pwdLastSet) { (Get-Date $pwdLastSet).ToString("dd.MM.yyyy HH:mm") } else { "Nie" }
            $pwdAgeDays = if ($pwdLastSet) { [int][math]::Round(((Get-Date) - $pwdLastSet).TotalDays, 0) } else { 9999 }

            $findings = [System.Collections.Generic.List[string]]::new()
            if ($isProtected -eq "JA") { $findings.Add("FEHLER: MSOL darf NICHT in Protected Users sein (bricht Sync ab)!") }
            if ($isNotDelegated -ne "JA") { $findings.Add("Flag 'Nicht delegierbar' setzen") }
            $finalFinding = if ($findings.Count -gt 0) { $findings -join " | " } else { "OK (Entra Sync Service)" }

            [void]$table.Rows.Add("Microsoft Entra Connect Sync", $user.SamAccountName, $user.Name, $(if($user.Enabled){"Aktiv"}else{"Deaktiviert"}), $pwdDateStr, $pwdAgeDays, $isProtected, $isNotDelegated, $finalFinding, $user.DistinguishedName)
        }
    } catch {}

    # 3. Entra Seamless SSO Computerkonto
    try {
        $ssoAccounts = Get-ADComputer -Filter {SamAccountName -like "AZUREADSSOACC*"} -Properties MemberOf, AccountNotDelegated, Enabled, SID, PasswordLastSet, Description -ErrorAction SilentlyContinue
        foreach ($comp in $ssoAccounts) {
            $groups = $comp.MemberOf | ForEach-Object { (Get-ADGroup -Identity $_).Name } | Sort-Object
            $isProtected = if ($groups -contains "Protected Users") { "JA" } else { "NEIN" }
            $isNotDelegated = if ($comp.AccountNotDelegated) { "JA" } else { "NEIN" }

            $pwdLastSet = $comp.PasswordLastSet
            $pwdDateStr = if ($pwdLastSet) { (Get-Date $pwdLastSet).ToString("dd.MM.yyyy HH:mm") } else { "Nie" }
            $pwdAgeDays = if ($pwdLastSet) { [int][math]::Round(((Get-Date) - $pwdLastSet).TotalDays, 0) } else { 9999 }

            $findings = [System.Collections.Generic.List[string]]::new()
            if ($pwdAgeDays -gt 90) { $findings.Add("Hinweis: Key-Rollover empfohlen (Schlüssel > 90 Tage)") }
            $finalFinding = if ($findings.Count -gt 0) { $findings -join " | " } else { "OK (Seamless SSO Secret)" }

            [void]$table.Rows.Add("Entra Seamless SSO Kerberos", $comp.SamAccountName, $comp.Name, $(if($comp.Enabled){"Aktiv"}else{"Deaktiviert"}), $pwdDateStr, $pwdAgeDays, $isProtected, $isNotDelegated, $finalFinding, $comp.DistinguishedName)
        }
    } catch {}

    # 4. Group Managed Service Accounts (gMSA)
    try {
        $gmsaAccounts = Get-ADServiceAccount -Filter * -Properties MemberOf, AccountNotDelegated, Enabled, SID, PasswordLastSet, Description -ErrorAction SilentlyContinue
        foreach ($sa in $gmsaAccounts) {
            $groups = $sa.MemberOf | ForEach-Object { (Get-ADGroup -Identity $_).Name } | Sort-Object
            $isProtected = if ($groups -contains "Protected Users") { "JA" } else { "NEIN" }
            $isNotDelegated = if ($sa.AccountNotDelegated) { "JA" } else { "NEIN" }

            $pwdLastSet = $sa.PasswordLastSet
            $pwdDateStr = if ($pwdLastSet) { (Get-Date $pwdLastSet).ToString("dd.MM.yyyy HH:mm") } else { "Automatisch" }
            $pwdAgeDays = if ($pwdLastSet) { [int][math]::Round(((Get-Date) - $pwdLastSet).TotalDays, 0) } else { 0 }

            [void]$table.Rows.Add("Group Managed Service Account (gMSA)", $sa.SamAccountName, $sa.Name, $(if($sa.Enabled){"Aktiv"}else{"Deaktiviert"}), $pwdDateStr, $pwdAgeDays, $isProtected, $isNotDelegated, "OK (AD Managed Rotation)", $sa.DistinguishedName)
        }
    } catch {}

    $gridServices.DataSource = $table
    $gridServices.Columns["DistinguishedName"].Visible = $false
})

$gridServices.Add_SelectionChanged({
    if ($gridServices.SelectedRows.Count -gt 0) {
        $row = $gridServices.SelectedRows[0]
        $selectedUser = [string]$row.Cells["SamAccountName"].Value
        $selectedRole = [string]$row.Cells["Dienst-Funktion"].Value
        $selectedStatus = [string]$row.Cells["Status (Aktiv)"].Value
        $selectedPwdDate = [string]$row.Cells["Schlüssel/PW geändert"].Value
        $selectedPwdAge = [int]$row.Cells["Alter (Tage)"].Value
        $selectedProtected = [string]$row.Cells["Protected Users?"].Value
        $selectedNotDelegated = [string]$row.Cells["Nicht Delegierbar"].Value
        $selectedFinding = [string]$row.Cells["Empfehlung / Audit-Befund"].Value
        $selectedDn = [string]$row.Cells["DistinguishedName"].Value

        $lblSvcGroups.Text = "2. Gruppenmitgliedschaften des Dienstkontos: [$selectedUser]"
        $lblSvcDesc.Text = "3. Beschreibung & Sicherheitsleitfaden: [$selectedUser]"

        $txtSvcDesc.Text = Get-ServiceAccountKnowledgeBaseText -userName $selectedUser -role $selectedRole -status $selectedStatus -pwdDateStr $selectedPwdDate -pwdAgeDays $selectedPwdAge -isProtected $selectedProtected -isNotDelegated $selectedNotDelegated -findings $selectedFinding -dn $selectedDn

        $tableSvcGroups = New-Object System.Data.DataTable
        [void]$tableSvcGroups.Columns.Add("Gruppen-Name")
        [void]$tableSvcGroups.Columns.Add("Herkunft / Typ")
        [void]$tableSvcGroups.Columns.Add("Erstellungsdatum")
        [void]$tableSvcGroups.Columns.Add("SID")
        [void]$tableSvcGroups.Columns.Add("Beschreibung")

        $domainAdminsGroup = Get-ResolvedADGroup "$domainSid-512"
        $daCreationDate = if ($domainAdminsGroup.whenCreated) { (Get-Date $domainAdminsGroup.whenCreated).Date } else { $null }

        try {
            $userObj = Get-ADObject -Identity $selectedDn -Properties memberOf -ErrorAction Stop
            $allGroupDns = [System.Collections.Generic.List[string]]::new()
            if ($userObj.memberOf) {
                foreach ($g in $userObj.memberOf) { $allGroupDns.Add($g) }
            }

            if ($allGroupDns.Count -eq 0) {
                [void]$tableSvcGroups.Rows.Add("-- Keine direkten Gruppen --", "-", "-", "-", "-")
            } else {
                foreach ($grpDn in ($allGroupDns | Sort-Object)) {
                    try {
                        $grpObj = Get-ADGroup -Identity $grpDn -Properties Description, whenCreated, SID -ErrorAction Stop
                        $origin = Get-GroupOriginType -grpSid $grpObj.SID.Value -grpCreated $grpObj.whenCreated -refDate $daCreationDate
                        $createdStr = if ($grpObj.whenCreated) { (Get-Date $grpObj.whenCreated).ToString("dd.MM.yyyy HH:mm") } else { "-" }

                        [void]$tableSvcGroups.Rows.Add(
                            $grpObj.Name,
                            $origin,
                            $createdStr,
                            $grpObj.SID.Value,
                            $grpObj.Description
                        )
                    } catch {
                        [void]$tableSvcGroups.Rows.Add($grpDn, "Unbekannt", "-", "-", "Gruppe konnte nicht aufgelöst werden")
                    }
                }
            }
        } catch {
            [void]$tableSvcGroups.Rows.Add("Fehler", "Gruppen konnten nicht geladen werden", "-", "-", "-")
        }

        $gridSvcGroups.DataSource = $tableSvcGroups
    }
})

# =========================================================================
# Event: TAB 3 - Privilegierte Gruppen Scan
# =========================================================================
$loadGroupsAction = {
    $gridGroups.DataSource = $null
    $tableGroups = New-Object System.Data.DataTable
    [void]$tableGroups.Columns.Add("Gruppe")
    [void]$tableGroups.Columns.Add("Herkunft / Typ")
    [void]$tableGroups.Columns.Add("Erstellungsdatum")
    [void]$tableGroups.Columns.Add("Anzahl Mitglieder", [int])
    [void]$tableGroups.Columns.Add("Sicherheits-Bewertung")
    [void]$tableGroups.Columns.Add("SID")
    [void]$tableGroups.Columns.Add("DistinguishedName")
    [void]$tableGroups.Columns.Add("Description")

    $domainAdminsGroup = Get-ResolvedADGroup "$domainSid-512"
    $daCreationDate = if ($domainAdminsGroup.whenCreated) { (Get-Date $domainAdminsGroup.whenCreated).Date } else { $null }

    $processedGroups = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($id in $privilegedGroupIdentifiers) {
        $grp = Get-ResolvedADGroup $id
        if ($grp -and -not $processedGroups.Contains($grp.DistinguishedName)) {
            [void]$processedGroups.Add($grp.DistinguishedName)

            $memberCount = 0
            try {
                $members = Get-ADGroupMember -Identity $grp.DistinguishedName -ErrorAction Stop
                $memberCount = @($members).Count
            } catch {
                if ($grp.Members) { $memberCount = @($grp.Members).Count }
            }

            $originType = Get-GroupOriginType -grpSid $grp.SID.Value -grpCreated $grp.whenCreated -refDate $daCreationDate
            $createdStr = if ($grp.whenCreated) { (Get-Date $grp.whenCreated).ToString("dd.MM.yyyy HH:mm") } else { "-" }

            $sidVal = $grp.SID.Value
            $isSchemaAdmin = $sidVal.EndsWith("-518")
            $isEnterpriseAdmin = $sidVal.EndsWith("-519")

            $bewertung = "OK"
            if ($isSchemaAdmin) {
                if ($memberCount -gt 0) {
                    $bewertung = "CRITICAL: Schema-Admins sollte im laufenden Betrieb LEER sein!"
                } else {
                    $bewertung = "Optimal (Gruppe ist leer)"
                }
            }
            elseif ($isEnterpriseAdmin) {
                if ($memberCount -gt 0) {
                    $bewertung = "Hinweis: Organisations-Admins sollte im Alltag leer sein"
                } else {
                    $bewertung = "Optimal (Gruppe ist leer)"
                }
            }

            [void]$tableGroups.Rows.Add(
                $grp.Name,
                $originType,
                $createdStr,
                $memberCount,
                $bewertung,
                $grp.SID.Value,
                $grp.DistinguishedName,
                $grp.Description
            )
        }
    }

    $gridGroups.DataSource = $tableGroups
    $gridGroups.Columns["DistinguishedName"].Visible = $false
    $gridGroups.Columns["SID"].Visible = $false
    $gridGroups.Columns["Description"].Visible = $false
}

$btnScanGroups.Add_Click($loadGroupsAction)

$gridGroups.Add_SelectionChanged({
    if ($gridGroups.SelectedRows.Count -gt 0) {
        $selectedName = $gridGroups.SelectedRows[0].Cells["Gruppe"].Value
        $selectedOrigin = $gridGroups.SelectedRows[0].Cells["Herkunft / Typ"].Value
        $selectedCreated = $gridGroups.SelectedRows[0].Cells["Erstellungsdatum"].Value
        $selectedDn = $gridGroups.SelectedRows[0].Cells["DistinguishedName"].Value
        $selectedSid = $gridGroups.SelectedRows[0].Cells["SID"].Value
        $selectedDesc = $gridGroups.SelectedRows[0].Cells["Description"].Value

        $lblMembers.Text = "2. Mitglieder der ausgewählten Gruppe: [$selectedName]"
        $lblDesc.Text = "3. Beschreibung & Sicherheits-Erklärung: [$selectedName]"

        $txtDescription.Text = Get-GroupKnowledgeBaseText -grpName $selectedName -sid $selectedSid -adDescription $selectedDesc -createdDate $selectedCreated -originType $selectedOrigin

        $tableMembers = New-Object System.Data.DataTable
        [void]$tableMembers.Columns.Add("SamAccountName")
        [void]$tableMembers.Columns.Add("Name")
        [void]$tableMembers.Columns.Add("Objekttyp")
        [void]$tableMembers.Columns.Add("Aktiviert")
        [void]$tableMembers.Columns.Add("DistinguishedName")

        try {
            $members = Get-ADGroupMember -Identity $selectedDn -ErrorAction Stop
            if (@($members).Count -eq 0) {
                [void]$tableMembers.Rows.Add("-- Keine --", "-- Gruppe ist leer --", "-", "-", "-")
            } else {
                foreach ($m in $members) {
                    $enabledStatus = "-"
                    if ($m.objectClass -eq "user") {
                        try {
                            $uObj = Get-ADUser -Identity $m.distinguishedName -Properties Enabled
                            $enabledStatus = if ($uObj.Enabled) { "True" } else { "False" }
                        } catch { $enabledStatus = "?" }
                    }
                    [void]$tableMembers.Rows.Add($m.SamAccountName, $m.name, $m.objectClass, $enabledStatus, $m.distinguishedName)
                }
            }
        } catch {
            [void]$tableMembers.Rows.Add("Fehler", "Mitglieder konnten nicht geladen werden", "-", "-", "-")
        }
        $gridMembers.DataSource = $tableMembers
    }
})

# =========================================================================
# Startausführung beim Laden des Fensters
# =========================================================================
$form.Add_Shown({
    $btnScanAdmins.PerformClick()
    $btnScanServices.PerformClick()
    & $loadGroupsAction
})

[void]$form.ShowDialog()