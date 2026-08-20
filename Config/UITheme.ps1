<#
================================================================================
 CONFIG: UI THEME & TABELLEN-LAYOUT
================================================================================
#>
$script:UITheme = @{
    # --- Tabellen / DataGridView Layout ---
    HeaderHeight          = 36          # Höhe der Spaltenüberschriften in Pixel
    RowHeight             = 28          # Höhe jeder Datenzeile in Pixel
    HeaderPaddingLeft     = 8           # Innenabstand links im Header
    HeaderPaddingRight    = 8           # Innenabstand rechts im Header
    CellPaddingLeft       = 8           # Innenabstand links in den Zellen
    CellPaddingRight      = 8           # Innenabstand rechts in den Zellen
    CellPaddingTop        = 2           # Innenabstand oben in den Zellen
    CellPaddingBottom     = 2           # Innenabstand unten in den Zellen
    
    # --- Schriftarten ---
    FontFamily            = "Segoe UI"  # Standard-Schriftart
    HeaderFontSize        = 9.5         # Schriftgröße Header (pt)
    CellFontSize          = 9.0         # Schriftgröße Tabelleninhalt (pt)
    HeaderFontStyle       = [System.Drawing.FontStyle]::Bold
    CellFontStyle         = [System.Drawing.FontStyle]::Regular
    
    # --- Farbpalette Header & Tabellen ---
    HeaderBackColor       = [System.Drawing.Color]::FromArgb(238, 242, 246)
    HeaderForeColor       = [System.Drawing.Color]::FromArgb(40, 40, 40)
    GridLineColor         = [System.Drawing.Color]::FromArgb(226, 232, 240)
    RowBackColor          = [System.Drawing.Color]::White
    RowAltBackColor       = [System.Drawing.Color]::FromArgb(250, 252, 254)
    SelectionBackColor    = [System.Drawing.Color]::FromArgb(203, 228, 249)
    SelectionForeColor    = [System.Drawing.Color]::Black

    # --- Akzent-Farben ---
    AccentColor           = [System.Drawing.Color]::FromArgb(0, 120, 215)
    AccentColorDark       = [System.Drawing.Color]::FromArgb(24, 37, 55)

    # --- Standard-Fensterabmessungen ---
    DefaultToolWidth      = 1200
    DefaultToolHeight     = 780
    HeaderPanelHeight     = 60
}
