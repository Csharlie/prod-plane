' modConfig.bas - Globális beállítások és konstansok
' Ez a modul tartalmazza az alkalmazás konfigurációs beállításait

Option Explicit

' Fájl elérési útvonalak
Public Const PART_DETAILS_PATH As String = "P:\_Departments\PW\10. Transfer\Sárdy Péter\excel\pw-plan\part-details.xlsm"
Public Const DEFAULT_SAVE_DIR As String = "\OneDrive - Mercedes-Benz (corpdir.onmicrosoft.com)\Documents\pw-plan\"

' Színkódok
Public Const COLOR_HEADER_BG As Long = RGB(26, 26, 26)
Public Const COLOR_HEADER_FONT As Long = RGB(242, 242, 242)
Public Const COLOR_ALTERNATING_1 As Long = RGB(255, 255, 255)
Public Const COLOR_ALTERNATING_2 As Long = RGB(220, 220, 220)
Public Const COLOR_SPECIAL_ROW As Long = RGB(26, 26, 26)
Public Const COLOR_SPECIAL_FONT As Long = RGB(242, 242, 242)
Public Const COLOR_BORDER As Long = RGB(191, 191, 191)

' Táblázat beállítások
Public Const TABLE_NAME As String = "PwPlan"
Public Const TABLE_STYLE As String = "TableStyleMedium9"

' Oszlopnevek konstansok
Public Const COL_ORDER As String = "Auftrag"
Public Const COL_MATERIAL As String = "Material"
Public Const COL_MATERIAL_TEXT As String = "Materialkurztext"
Public Const COL_DRAWING_GEOMETRY As String = "Zeichnungsgeometriestand"
Public Const COL_CHARGE As String = "Charge"
Public Const COL_AMOUNT As String = "Vorgangsmenge (MEINH)"
Public Const COL_LT1 As String = "LT1-Bedarf"
Public Const COL_ZGS As String = "ZGS"
Public Const COL_ABPR As String = "Abpr"
Public Const COL_MENGE As String = "Menge"
Public Const COL_LT As String = "LT"
Public Const COL_WORKPLACE As String = "Arbeitsplatz"
Public Const COL_START_DATE As String = "Spätestes Startdatum"

' Elrejtendő oszlopok
Public Function GetColumnsToHide() As Variant
    GetColumnsToHide = Array( _
        "Arbeitsplatz", _
        "Rückgem. Gutmenge (MEINH)", _
        "Spätestes Startdatum", _
        "Späteste Startzeit", _
        "Systemstatus 1", _
        "LT1-Soll-Füllmenge", _
        "Mengeneinheit Vrg. (=MEINH)", _
        "Bestandsreichweite Werk", _
        "Bestandsreichweite Werk (MEINH)", _
        "Ist-Reichweite", _
        "Ist-Reichweite (MEINH)", _
        "Datum des Iststarts", _
        "Zeit des Iststarts", _
        "Datum des Istendes", _
        "Datum des Istbeendens", _
        "Zeit des Istbeendens", _
        "Anwenderstatus", _
        "LT1-Bestand", _
        "Datum des Iststarts2", _
        "Ladungsträger 1", _
        "Ladungsträger 2", _
        "Systemstatus", _
        "Vorgang" _
    )
End Function

' Szöveg cserék
Public Function GetTextReplacements() As Object
    Dim replacements As Object
    Set replacements = CreateObject("Scripting.Dictionary")
    
    replacements.Add "2024_", ""
    replacements.Add "2025_", ""
    replacements.Add "BEPLANKUNG", "BPL"
    replacements.Add "INNENTEIL", "INT"
    
    Set GetTextReplacements = replacements
End Function

' Oszlop fejléc átnevezések
Public Function GetHeaderReplacements() As Object
    Dim replacements As Object
    Set replacements = CreateObject("Scripting.Dictionary")
    
    replacements.Add COL_DRAWING_GEOMETRY, "ZGS"
    replacements.Add COL_CHARGE, "Abpr"
    replacements.Add COL_AMOUNT, "Menge"
    replacements.Add COL_LT1, "LT"
    
    Set GetHeaderReplacements = replacements
End Function

' Oszlop szélesség beállítások
Public Function GetColumnWidths() As Object
    Dim widths As Object
    Set widths = CreateObject("Scripting.Dictionary")
    
    widths.Add "Auftrag", 8
    widths.Add "Material", 14
    widths.Add "Materialkurztext", 24
    widths.Add "ZGS", 4
    widths.Add "Abpr", 6
    widths.Add "Menge", 7
    widths.Add "Kurztext Vorgang", 28
    widths.Add "LT", 4
    
    Set GetColumnWidths = widths
End Function

' Oszlop igazítás beállítások
Public Function GetColumnAlignments() As Object
    Dim alignments As Object
    Set alignments = CreateObject("Scripting.Dictionary")
    
    ' Középre igazított oszlopok
    alignments.Add "ZGS", xlCenter
    alignments.Add "Abpr", xlCenter
    
    ' Jobbra igazított oszlopok
    alignments.Add "Menge", xlRight
    alignments.Add "LT", xlRight
    
    Set GetColumnAlignments = alignments
End Function