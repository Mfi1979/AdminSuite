<#
================================================================================
 CONFIG: I18N SPRACHVERWALTUNG (REIN DATEIBASIERT)
================================================================================
#>

$global:CurrentLang = "DE"
$global:I18N = @{}

# Funktion: Registriert eine geladene JSON-Sprachdatei (egal ob lokal oder aus dem Web)
function Register-LanguageJson {
    param(
        [string]$FileName,
        [string]$JsonContent
    )
    try {
        $jsonObj = $JsonContent | ConvertFrom-Json
        # Ermittelt z. B. aus "de-DE.json" das Kürzel "DE"
        $langCode = ($FileName.Split('.')[0].Split('-')[0]).ToUpper()

        if (-not $global:I18N.ContainsKey($langCode)) {
            $global:I18N[$langCode] = @{}
        }

        # Eigenschaften als Key-Value in die Hashtable eintragen
        $jsonObj.PSObject.Properties | ForEach-Object {
            $global:I18N[$langCode][$_.Name] = $_.Value
        }
    } catch {
        Write-Warning "Fehler beim Verarbeiten der Sprachdatei: $FileName"
    }
}

# Funktion: Text holen mit Fallback
function Get-Text {
    param([string]$Key)

    # 1. Gewünschte Sprache prüfen
    if ($global:I18N.ContainsKey($global:CurrentLang)) {
        if ($global:I18N[$global:CurrentLang].ContainsKey($Key)) {
            return $global:I18N[$global:CurrentLang][$Key]
        }
    }

    # 2. Fallback auf Deutsch (DE)
    if ($global:I18N.ContainsKey("DE")) {
        if ($global:I18N["DE"].ContainsKey($Key)) {
            return $global:I18N["DE"][$Key]
        }
    }

    # 3. Wenn nirgends gefunden: Schlüsselnamen selbst zurückgeben
    return $Key
}