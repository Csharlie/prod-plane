Attribute VB_Name = "modCalculations"
' modCalculations.bas - Sz�m�t�sokkal kapcsolatos funkci�k

Option Explicit

Public Sub CalculateSums()
    ' SumMinMenge elj�r�s h�v�sa
    SumMinMenge
End Sub

Sub SumMinMenge()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' "Auftrag" és "Menge" oszlopok keresése a PwPlan táblázatban
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("PwPlan")
    
    Dim orderCol As Long, mengeCol As Long
    orderCol = GetColumnByName(tbl, "Auftrag")
    mengeCol = GetColumnByName(tbl, "Menge")
    
    If orderCol = 0 Or mengeCol = 0 Then
        MsgBox "A szükséges oszlopok nem találhatók!", vbExclamation
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = tbl.Range.Rows.Count
    Dim startRow As Long: startRow = 2 ' Kezdő sor
    
    ' Külső ciklus a szakaszok feldolgozásához
    Do While startRow <= lastRow
        ' Szótár létrehozása az aktuális szakasz minimum értékeihez
        Dim orderMengeDict As Object
        Set orderMengeDict = CreateObject("Scripting.Dictionary")
        
        ' Megkeressük a szakasz végét (következő üres cella)
        Dim endRow As Long: endRow = startRow
        Do While endRow <= lastRow
            If IsEmpty(tbl.Range.Cells(endRow, mengeCol).Value) Then
                Exit Do
            End If
            endRow = endRow + 1
        Loop
        
        ' Ha nincs több adat, kilépünk
        If endRow <= startRow Then
            Exit Do
        End If
        
        ' Minimum értékek összegyűjtése a szakaszban
        Dim i As Long
        For i = startRow To endRow - 1
            Dim currentOrder As String
            currentOrder = tbl.Range.Cells(i, orderCol).Value
            
            If currentOrder <> "" Then
                Dim mengeValue As Variant
                mengeValue = tbl.Range.Cells(i, mengeCol).Value
                
                If IsNumeric(mengeValue) Then
                    If Not orderMengeDict.Exists(currentOrder) Then
                        orderMengeDict.Add currentOrder, mengeValue
                    ElseIf mengeValue < orderMengeDict(currentOrder) Then
                        orderMengeDict(currentOrder) = mengeValue
                    End If
                End If
            End If
        Next i
        
        ' SZUM képlet létrehozása az aktuális szakaszhoz
        If orderMengeDict.Count > 0 Then
            Dim sumFormula As String
            sumFormula = "=SUM("
            
            ' Minimum értékek celláinak összegyűjtése
            Dim order As Variant
            For Each order In orderMengeDict.Keys
                For i = startRow To endRow - 1
                    If tbl.Range.Cells(i, orderCol).Value = order Then
                        If tbl.Range.Cells(i, mengeCol).Value = orderMengeDict(order) Then
                            sumFormula = sumFormula & tbl.Range.Cells(i, mengeCol).Address(False, False) & ","
                            Exit For
                        End If
                    End If
                Next i
            Next order
            
            ' Képlet befejezése és beszúrása
            If Right(sumFormula, 1) = "," Then
                sumFormula = Left(sumFormula, Len(sumFormula) - 1)
            End If
            sumFormula = sumFormula & ")"
            tbl.Range.Cells(endRow, mengeCol).Formula = sumFormula
        End If
        
        ' Következő szakasz kezdőpontja
        startRow = endRow + 1
    Loop
End Sub
