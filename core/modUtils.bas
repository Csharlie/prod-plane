' modUtils.bas - Általános segédfunkciók
' Ez a modul általános segédfunkciókat tartalmaz, amelyeket több modul is használhat

Option Explicit

' Oszlop keresése a fejléc neve alapján
Public Function FindColumn(ws As Worksheet, headerName As String) As Long
    Dim lastCol As Long
    Dim col As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If ws.Cells(1, col).Value = headerName Then
            FindColumn = col
            Exit Function
        End If
    Next col
    
    FindColumn = 0 ' Nem található
End Function

' Oszlop indexének megkeresése név alapján egy táblázatban
Public Function GetColumnByName(tbl As ListObject, colName As String) As Long
    Dim col As ListColumn
    
    For Each col In tbl.ListColumns
        If col.Name = colName Then
            GetColumnByName = col.Index
            Exit Function
        End If
    Next col
    
    GetColumnByName = 0 ' Nem található
End Function

' SZUM képlethez cellatartomány szöveg összeállítása
Public Function GetSumRangeString(cells As Collection) As String
    Dim result As String
    Dim i As Long
    
    For i = 1 To cells.Count
        result = result & cells(i)
        If i < cells.Count Then result = result & ","
    Next i
    
    GetSumRangeString = result
End Function

' Ellenőrzi, hogy egy fájl létezik-e
Public Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

' Ellenőrzi, hogy egy mappa létezik-e
Public Function FolderExists(folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

' Létrehoz egy mappát, ha még nem létezik
Public Sub CreateFolderIfNotExists(folderPath As String)
    If Not FolderExists(folderPath) Then
        MkDir folderPath
    End If
End Sub

' Felhasználói név lekérése a környezeti változókból
Public Function GetUserName() As String
    GetUserName = Environ$("USERNAME")
End Function

' Aktuális dátum és idő formázott szövegként
Public Function GetFormattedDateTime() As String
    GetFormattedDateTime = Format(Now, "yyyy.mm.dd. hh:mm")
End Function

' Fájl utolsó módosításának dátuma és ideje
Public Function GetLastModifiedDateTime(filePath As String) As String
    On Error Resume Next
    GetLastModifiedDateTime = Format(FileDateTime(filePath), "yyyy.mm.dd. hh:mm")
    On Error GoTo 0
    
    If Err.Number <> 0 Then
        GetLastModifiedDateTime = ""
    End If
End Function

' Debug üzenet kiírása az Immediate ablakba
Public Sub DebugPrint(message As String)
    Debug.Print Format(Now, "hh:mm:ss") & " - " & message
End Sub