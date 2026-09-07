<#
================================================================================
 CONFIG: I18N DYNAMISCHER JSON-LOADER & GET-TEXT FUNKTION
================================================================================
#>

$script:CurrentLang = "DE"
$script:I18N = @{}

# Sprach-Mapping von Kurzcode auf Dateinamen
$script:LangFileMap = @{
    "DE" = "de-DE.json"
    "EN" = "en-US.json"
}

# Basisverzeichnis des Projekts ermitteln
$scriptRoot = Split-Path -Parent $PSScriptRoot
$langFolder = Join-Path $scriptRoot "Languages"
if (-not (Test-Path $langFolder)) {
    # Fallback falls direkt im Root ausgeführt
    $langFolder = Join-Path $PSScriptRoot "Languages"
}

# Alle verfügbaren Sprachdateien einlesen
function Initialize-Languages {
    if (Test-Path $langFolder) {
        Get-ChildItem -Path $langFolder -Filter "*.json" | ForEach-Object {
            try {
                $content = Get-Content -Path $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                
                # Datei z.B. "de-DE.json" -> Key "DE", "en-US.json" -> Key "EN"
                $shortKey = $_.BaseName.Split('-')[0].ToUpper()
                
                # In Hashtable konvertieren
                $hash = @{}
                $content.PSObject.Properties | ForEach-Object {
                    $hash[$_.Name] = $_.Value
                }
                
                $script:I18N[$shortKey] = $hash
            } catch {
                Write-Warning "Fehler beim Laden der Sprachdatei: $($_.FullName)"
            }
        }
    }
}

# Text-Abruf mit Fallback
function Get-Text {
    param([string]$Key)
    
    if ($script:I18N.ContainsKey($script:CurrentLang) -and $script:I18N[$script:CurrentLang].ContainsKey($Key)) {
        return $script:I18N[$script:CurrentLang][$Key]
    }
    # Fallback auf DE falls in Zielsprache nicht gepflegt
    if ($script:I18N.ContainsKey("DE") -and $script:I18N["DE"].ContainsKey($Key)) {
        return $script:I18N["DE"][$Key]
    }
    return $Key
}

# Sprache wechseln
function Set-Language {
    param([string]$LangCode)
    if ($script:I18N.ContainsKey($LangCode.ToUpper())) {
        $script:CurrentLang = $LangCode.ToUpper()
    }
}

# Beim Laden direkt initialisieren
Initialize-Languages