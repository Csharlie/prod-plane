' modFormatter.bas - Formázással kapcsolatos funkciók
' Ez a modul felelős a munkafüzet formázásáért

Option Explicit

' Formázás végrehajtása
Public Sub FormatWorksheet()
    ' Aktív munkalap
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Oszlopok és sorok meghatározása
    Dim lastRow As Long, lastCol As Long
    Dim orderCol As Long, materialCol As Long, materialTextCol As Long
    Dim arbeitsplatzCol As Long, spatestesStartdatumCol As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, ws.Columns.Count).End(xlUp).Row
    
    orderCol = modUtils.FindColumn(ws, modConfig.COL_ORDER)
    materialCol = modUtils.FindColumn(ws, modConfig.COL_MATERIAL)
    materialTextCol = modUtils.FindColumn(ws, modConfig.COL_MATERIAL_TEXT)
    arbeitsplatzCol = modUtils.FindColumn(ws, modConfig.COL_WORKPLACE)
    spatestesStartdatumCol = modUtils.FindColumn(ws, modConfig.COL_START_DATE)
    
    ' Sorok formázása
    FormatRows ws, orderCol, materialTextCol, lastRow, lastCol
    
    ' Táblázat formázása
    FormatTable ws, lastRow, lastCol
    
    ' Oszlopok elrejtése
    HideColumns ws, lastCol
    
    ' Speciális munkahelyek kezelése
    HandleSpecialWorkplaces ws, orderCol, arbeitsplatzCol, spatestesStartdatumCol, lastRow
    
    ' Oszlopok szélességének beállítása
    SetColumnWidths ws
    
    ' Oszlopok igazításának beállítása
    SetColumnAlignments ws
    
    ' Oldalbeállítások konfigurálása
    ConfigurePageSetup ws, spatestesStartdatumCol, lastRow
End Sub

' Sorok formázása az "Auftrag" oszlop értékei szerint
Private Sub FormatRows(ws As Worksheet, orderCol As Long, materialTextCol As Long, lastRow As Long, lastCol As Long)
    Dim currentColor As Long
    Dim previousOrder As String, currentOrder As String
    Dim i As Long
    
    currentColor = modConfig.COLOR_ALTERNATING_1 ' Kezdeti szín: fehér
    previousOrder = "" ' Kezdetben üres
    
    For i = 2 To lastRow
        If Not ws.Rows(i).Hidden Then
            currentOrder = ws.Cells(i, orderCol).Value
            Dim rowEndCol As Long
            rowEndCol = ws.Cells(i, ws.Columns.Count).End(xlToLeft).Column
            
            ' Ellenőrizzük, hogy "Instandhaltungsauftrag" sor-e
            If ws.Cells(i, materialTextCol).Value = "Instandhaltungsauftrag" Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, rowEndCol)).Interior.Color = modConfig.COLOR_SPECIAL_ROW
                ws.Range(ws.Cells(i, 1), ws.Cells(i, rowEndCol)).Font.Color = modConfig.COLOR_SPECIAL_FONT
                
                ' Speciális mezők törlése
                ClearSpecialFields ws, i
            Else
                ' Normál váltakozó színek
                If currentOrder = previousOrder Then
                    ws.Range(ws.Cells(i, 1), ws.Cells(i, rowEndCol)).Interior.Color = currentColor
                Else
                    If currentColor = modConfig.COLOR_ALTERNATING_1 Then
                        currentColor = modConfig.COLOR_ALTERNATING_2
                    Else
                        currentColor = modConfig.COLOR_ALTERNATING_1
                    End If
                    ws.Range(ws.Cells(i, 1), ws.Cells(i, rowEndCol)).Interior.Color = currentColor
                End If
            End If
            
            previousOrder = currentOrder
        End If
    Next i
End Sub

' Speciális mezők törlése
Private Sub ClearSpecialFields(ws As Worksheet, rowIndex As Long)
    Dim zgsCol As Long, mengeCol As Long, ltCol As Long
    
    ' Oszlopok keresése
    On Error Resume Next
    zgsCol = modUtils.FindColumn(ws, "ZGS")
    mengeCol = modUtils.FindColumn(ws, "Menge")
    ltCol = modUtils.FindColumn(ws, "LT")
    On Error GoTo 0
    
    ' ZGS oszlopban '0' érték törlése
    If zgsCol > 0 Then
        If ws.Cells(rowIndex, zgsCol).Value = "0" Then
            ws.Cells(rowIndex, zgsCol).ClearContents
        End If
    End If
    
    ' Menge oszlopban '100' érték törlése
    If mengeCol > 0 Then
        If ws.Cells(rowIndex, mengeCol).Value = "100" Then
            ws.Cells(rowIndex, mengeCol).ClearContents
        End If
    End If
    
    ' LT oszlopban '0' érték törlése
    If ltCol > 0 Then
        If ws.Cells(rowIndex, ltCol).Value = "0" Then
            ws.Cells(rowIndex, ltCol).ClearContents
        End If
    End If
End Sub

' Táblázat formázása
Private Sub FormatTable(ws As Worksheet, lastRow As Long, lastCol As Long)
    Dim tbl As ListObject
    
    On Error Resume Next
    Set tbl = ws.ListObjects(modConfig.TABLE_NAME)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
        tbl.Name = modConfig.TABLE_NAME
        tbl.TableStyle = modConfig.TABLE_STYLE
    Else
        tbl.Resize ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    End If
    
    ' Automatikus szűrők eltávolítása
    tbl.ShowAutoFilter = False
    
    ' Szegély beállítása
    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    With tableRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = modConfig.COLOR_BORDER
    End With
    
    ' Fejléc formázása
    With tbl.HeaderRowRange
        .RowHeight = 25
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Color = modConfig.COLOR_HEADER_FONT
        .Interior.Color = modConfig.COLOR_HEADER_BG
    End With
    
    ' Adattartomány formázása
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
    
    With dataRange
        .RowHeight = 16
        .VerticalAlignment = xlCenter
    End With
End Sub

' Oszlopok elrejtése
Private Sub HideColumns(ws As Worksheet, lastCol As Long)
    Dim columnsToHide As Variant
    columnsToHide = modConfig.GetColumnsToHide()
    
    Dim col As Long
    For col = 1 To lastCol
        Dim headerValue As String
        headerValue = ws.Cells(1, col).Value
        
        Dim keywordIndex As Integer
        For keywordIndex = LBound(columnsToHide) To UBound(columnsToHide)
            If headerValue = columnsToHide(keywordIndex) Then
                ws.Columns(col).Hidden = True
                Exit For
            End If
        Next keywordIndex
    Next col
End Sub

' Speciális munkahelyek kezelése
Private Sub HandleSpecialWorkplaces(ws As Worksheet, orderCol As Long, arbeitsplatzCol As Long, spatestesStartdatumCol As Long, lastRow As Long)
    If arbeitsplatzCol > 0 Then
        If ws.Cells(2, arbeitsplatzCol).Value = "KT371041" Then
            ' Az "Auftrag" oszlop áthelyezése az első helyre
            If orderCol > 1 Then
                ws.Columns(orderCol).Cut
                ws.Columns(1).Insert Shift:=xlToRight
                Application.CutCopyMode = False
            End If
            
            ' Fejléc beállítás
            ws.PageSetup.CenterHeader = "&B" & "PW terv - Servo"
            ws.PageSetup.LeftHeader = "PW / Servo TL-XL 7.41" & Chr(10) & "Arbeitsplatz: KT371041"
            
        ElseIf ws.Cells(2, arbeitsplatzCol).Value = "KT371022" Then
            ' Rejtsük el az 'Auftrag' oszlopot
            If orderCol > 0 Then
                ws.Columns(orderCol).Hidden = True
            End If
            
            ' Rejtsük el az első oszlopot
            ws.Columns(1).Hidden = True
            
            ' Biztosítsuk, hogy a "Spätestes Startdatum" oszlop látható legyen
            If spatestesStartdatumCol > 0 Then
                ws.Columns(spatestesStartdatumCol).Hidden = False
            End If
        End If
        
        ' Most elrejtjük az 'Arbeitsplatz' oszlopot
        ws.Columns(arbeitsplatzCol).Hidden = True
    End If
End Sub

' Oszlopok szélességének beállítása
Private Sub SetColumnWidths(ws As Worksheet)
    Dim columnWidths As Object
    Set columnWidths = modConfig.GetColumnWidths()
    
    Dim header As Range
    For Each header In ws.Rows(1).Cells
        If columnWidths.Exists(header.Value) Then
            header.EntireColumn.ColumnWidth = columnWidths(header.Value)
        End If
    Next header
End Sub

' Oszlopok igazításának beállítása
Private Sub SetColumnAlignments(ws As Worksheet)
    Dim columnAlignments As Object
    Set columnAlignments = modConfig.GetColumnAlignments()
    
    Dim header As Range
    For Each header In ws.Rows(1).Cells
        If columnAlignments.Exists(header.Value) Then
            header.EntireColumn.HorizontalAlignment = columnAlignments(header.Value)
        End If
    Next header
End Sub

' Oldalbeállítások konfigurálása
Private Sub ConfigurePageSetup(ws As Worksheet, spatestesStartdatumCol As Long, lastRow As Long)
    With ws.PageSetup
        ' Margók
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(