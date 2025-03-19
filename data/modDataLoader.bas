' modDataLoader.bas - Adatok betöltése és előkészítése
' Ez a modul felelős az adatok betöltéséért és előkészítéséért

Option Explicit

' Part-details munkafüzet és táblázat referenciái
Private wbPartDetails As Workbook
Private wsPartDetails As Worksheet
Private tblPartDetails As ListObject

' Aktív munkalap és táblázat
Private ws As Worksheet
Private tbl As ListObject

' Oszlopindexek
Private orderCol As Long
Private materialCol As Long
Private idCol As Long
Private lastRow As Long
Private lastCol As Long

' Adatok betöltése és feldolgozása
Public Sub ProcessData()
    ' Referenciák inicializálása
    Set ws = ActiveSheet
    
    ' Oszlopok meghatározása
    InitializeColumns
    
    ' Part-details munkafüzet megnyitása
    OpenPartDetailsWorkbook
    
    ' ID oszlop beszúrása
    InsertIDColumn
    
    ' Part-details adatok előkészítése
    PreparePartDetailsData
    
    ' ID értékek feltöltése
    PopulateIDValues
    
    ' Part-details mentése
    wbPartDetails.Save
    
    ' Táblázat létrehozása
    CreateTable
    
    ' Fejlécek módosítása
    ModifyHeaders
    
    ' Szövegek cseréje
    ReplaceTextValues
End Sub

' Oszlopok inicializálása
Private Sub InitializeColumns()
    ' Utolsó oszlop meghatározása
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Rendelés és Anyag oszlopok keresése
    orderCol = modUtils.FindColumn(ws, modConfig.COL_ORDER)
    materialCol = modUtils.FindColumn(ws, modConfig.COL_MATERIAL)
    
    ' Oszlopok ellenőrzése
    modUtils.DebugPrint "Kezdeti Rendelés oszlop: " & orderCol
    modUtils.DebugPrint "Anyag oszlop: " & materialCol
    
    ' Utolsó sor meghatározása
    lastRow = ws.Cells(ws.Rows.Count, materialCol).End(xlUp).Row
End Sub

' ID oszlop beszúrása
Private Sub InsertIDColumn()
    ws.Columns(orderCol).Insert Shift:=xlToRight
    ws.Cells(1, orderCol).Value = "ID"
    orderCol = orderCol + 1 ' orderCol most már az eredeti oszlopra mutat, az ID oszlop előtte van
End Sub

' Part-details munkafüzet megnyitása
Private Sub OpenPartDetailsWorkbook()
    Set wbPartDetails = Workbooks.Open(modConfig.PART_DETAILS_PATH)
    Set wsPartDetails = wbPartDetails.Worksheets("ID DLR ST")
    Set tblPartDetails = wsPartDetails.ListObjects("ID_DLR_ST")
    
    ' Ellenőrizzük a part-details oszlopneveit
    modUtils.DebugPrint "Part-details oszlopok:"
    Dim colDetails As ListColumn
    For Each colDetails In tblPartDetails.ListColumns
        modUtils.DebugPrint colDetails.Name
    Next colDetails
End Sub

' Part-details adatok előkészítése
Private Sub PreparePartDetailsData()
    ' Oszlopok azonosítása
    Dim partNumberCol As Long
    partNumberCol = tblPartDetails.ListColumns("PartNumber_Plan").Index
    idCol = tblPartDetails.ListColumns("ID").Index
    
    ' PartNumber_String oszlop kezelése
    Dim stringCol As Long
    On Error Resume Next
    stringCol = tblPartDetails.ListColumns("PartNumber_String").Index
    On Error GoTo 0
    
    If stringCol = 0 Then
        ' Ha nem létezik a PartNumber_String oszlop, létrehozzuk
        Set colDetails = tblPartDetails.ListColumns.Add
        colDetails.Name = "PartNumber_String"
        stringCol = colDetails.Index
        modUtils.DebugPrint "PartNumber_String oszlop létrehozva"
    End If
    
    ' Part-details anyagszámok konvertálása és másolása
    With tblPartDetails.ListColumns("PartNumber_Plan").DataBodyRange
        .NumberFormat = "@"
        .Value = .Value
        
        ' Értékek másolása a String oszlopba
        Dim i As Long
        For i = 1 To .Rows.Count
            tblPartDetails.DataBodyRange.Cells(i, stringCol).Value = _
                CStr(Trim(.Cells(i).Value))
        Next i
    End With
    
    modUtils.DebugPrint "Adatok konvertálva és másolva a PartNumber_String oszlopba"
    
    ' Most már a PartNumber_String oszlopban keresünk
    partNumberCol = stringCol
    modUtils.DebugPrint "partNumberCol: " & partNumberCol
End Sub

' ID értékek feltöltése
Private Sub PopulateIDValues()
    Dim i As Long, j As Long
    
    For i = 2 To lastRow
        Dim matValue As String
        matValue = Trim(ws.Cells(i, materialCol).Value)
        
        Dim found As Boolean
        found = False
        
        For j = 1 To tblPartDetails.ListRows.Count
            If matValue = Trim(tblPartDetails.DataBodyRange.Cells(j, partNumberCol).Value) Then
                ws.Cells(i, orderCol - 1).Value = tblPartDetails.DataBodyRange.Cells(j, idCol).Value
                found = True
                Exit For
            End If
        Next j

        If Not found Then
            modUtils.DebugPrint "Nincs egyezés: " & matValue
        End If
    Next i
End Sub

' Táblázat létrehozása
Private Sub CreateTable()
    On Error Resume Next
    Set tbl = ws.ListObjects(modConfig.TABLE_NAME)
    If Err.Number <> 0 Then
        Err.Clear
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
        If Not tbl Is Nothing Then
            tbl.Name = modConfig.TABLE_NAME
        End If
    Else
        tbl.Resize ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    End If
    On Error GoTo 0
End Sub

' Fejlécek módosítása
Private Sub ModifyHeaders()
    Dim headerReplacements As Object
    Set headerReplacements = modConfig.GetHeaderReplacements()
    
    Dim col As Long
    For col = 1 To lastCol
        Dim headerValue As String
        headerValue = ws.Cells(1, col).Value
        
        If headerReplacements.Exists(headerValue) Then
            ws.Cells(1, col).Value = headerReplacements(headerValue)
        End If
    Next col
End Sub

' Szövegek cseréje
Private Sub ReplaceTextValues()
    Dim textReplacements As Object
    Set textReplacements = modConfig.GetTextReplacements()
    
    Dim cell As Range
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) = False Then
            Dim key As Variant
            For Each key In textReplacements.Keys
                cell.Value = Replace(cell.Value, key, textReplacements(key))
            Next key
        End If
    Next cell
End Sub